using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;

namespace OstMsgExport
{
    sealed class FolderInfo
    {
        public uint Nid; public string Name; public string Path; public int ContentCount;
        public List<uint> Listed = new List<uint>();     // messages in the folder's contents table
        public List<uint> Extra = new List<uint>();      // message records whose parent is this folder but are not listed
    }

    static class Program
    {
        const string Version = "1.0.0";

        static int Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            int rc;
            try { rc = Run(args); }
            catch (Exception e) when (e is InvalidDataException || e is IOException || e is UnauthorizedAccessException)
            {
                Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine("Error: " + e.Message); Console.ResetColor();
                rc = 2;
            }
            // Started by double-click or by dropping a file on the .exe: keep the window open to read the result.
            if (OwnsConsole() && !Console.IsInputRedirected) { Console.WriteLine(); Console.Write("Press Enter to close."); Console.ReadLine(); }
            return rc;
        }

        [System.Runtime.InteropServices.DllImport("kernel32.dll")] static extern uint GetConsoleProcessList(uint[] list, uint count);
        static bool OwnsConsole() { try { return GetConsoleProcessList(new uint[4], 4) <= 1; } catch { return false; } }


        static void Usage()
        {
            Console.WriteLine($"OstMsgExport {Version} - exports cached messages from an Outlook OST/PST file to .msg files.");
            Console.WriteLine();
            Console.WriteLine("  OstMsgExport.exe <file.ost> [--folder <name>] [--out <folder>] [--list-folders] [--yes]");
            Console.WriteLine();
            Console.WriteLine("  --folder <name>   Folder to export (default: Inbox). A name or a path such as \"Inbox/Receipts\". Use * for all.");
            Console.WriteLine("  --out <folder>    Where to write the .msg files (default: a new folder on the Desktop).");
            Console.WriteLine("  --list-folders    Lists the folders and their message counts, then stops.");
            Console.WriteLine("  --yes             Exports without asking after the analysis.");
            Console.WriteLine("  --probe           Development: dumps the low-level file layout.");
            Console.WriteLine();
            Console.WriteLine("The OST file is only read. Work on a copy of it, not on the file Outlook is using.");
        }

        static int Run(string[] args)
        {
            string file = null, folderArg = "Inbox", outDir = null; bool list = false, yes = false;
            for (int i = 0; i < args.Length; i++)
            {
                string a = args[i];
                if (a == "--probe" && i + 1 < args.Length) { Probe.Run(args[++i]); return 0; }
                else if (a == "--folder" && i + 1 < args.Length) folderArg = args[++i];
                else if (a == "--out" && i + 1 < args.Length) outDir = args[++i];
                else if (a == "--list-folders") list = true;
                else if (a == "--yes" || a == "-y") yes = true;
                else if (a == "--help" || a == "-h" || a == "/?") { Usage(); return 0; }
                else if (!a.StartsWith("--") && file == null) file = a.Trim('"');
                else { Console.WriteLine("Unknown option: " + a); Usage(); return 1; }
            }
            if (file == null)
            {
                Usage();
                Console.WriteLine();
                Console.Write("Path of the OST file to read: ");
                file = (Console.ReadLine() ?? "").Trim().Trim('"');
                if (file.Length == 0) return 1;
            }
            file = Path.GetFullPath(file);
            if (!File.Exists(file)) throw new IOException("File not found: " + file);

            Console.WriteLine($"OstMsgExport {Version}");
            Console.WriteLine($"File: {file} ({Size(new FileInfo(file).Length)})");
            if (Process.GetProcessesByName("OUTLOOK").Length > 0 && file.IndexOf(@"\Microsoft\Outlook\", StringComparison.OrdinalIgnoreCase) >= 0 && file.IndexOf(@"\Backup\", StringComparison.OrdinalIgnoreCase) < 0)
                Warn("Outlook is running and this may be the file it is using. For reliable results use a copy, or close Outlook.");

            var sw = Stopwatch.StartNew();
            using (var ost = new OstFile(file))
            {
                Console.WriteLine($"Format: {(ost.Is4K ? "Outlook 2013+ OST/PST (4K pages)" : "Outlook 2003+ Unicode PST/OST")}; {ost.Nodes.Count:N0} records, {ost.Blocks.Count:N0} blocks indexed.");
                var folders = LoadFolders(ost);

                if (list)
                {
                    Console.WriteLine();
                    Console.WriteLine($"{"Messages",9} {"Extra",6}  Folder");
                    foreach (var f in folders.OrderBy(f => f.Path, StringComparer.OrdinalIgnoreCase))
                        Console.WriteLine($"{f.Listed.Count,9:N0} {(f.Extra.Count > 0 ? f.Extra.Count.ToString("N0") : ""),6}  {f.Path}");
                    Console.WriteLine();
                    Console.WriteLine("Extra = message records that belong to the folder but are no longer in its message list.");
                    return 0;
                }

                var chosen = SelectFolders(folders, folderArg);
                if (chosen.Count == 0)
                {
                    Warn($"No folder named \"{folderArg}\" was found. Folders in this file:");
                    foreach (var f in folders.OrderBy(f => f.Path, StringComparer.OrdinalIgnoreCase)) Console.WriteLine("   " + f.Path);
                    return 3;
                }

                // ---------- analysis: build every message in memory to get the exact count and size ----------
                Console.WriteLine();
                Console.WriteLine("Analysing (nothing is written yet)...");
                var work = new List<KeyValuePair<FolderInfo, uint>>();
                foreach (var f in chosen) { foreach (var n in f.Listed) work.Add(new KeyValuePair<FolderInfo, uint>(f, n)); foreach (var n in f.Extra) work.Add(new KeyValuePair<FolderInfo, uint>(f, n)); }
                long total = 0, onDisk = 0; int ok = 0, failed = 0, missing = 0, withWarnings = 0;
                var failures = new List<string>();
                for (int i = 0; i < work.Count; i++)
                {
                    if (i % 50 == 0) Console.Write($"\r   {i:N0} of {work.Count:N0}");
                    if (!ost.Nodes.TryGetValue(work[i].Value, out var e)) { missing++; continue; }
                    try
                    {
                        var m = MessageReader.Read(new Node(ost, e));
                        var bytes = MsgBuilder.Build(m);
                        ok++; total += bytes.Length; onDisk += (bytes.Length + 4095) / 4096 * 4096;
                        if (m.Warnings.Count > 0) withWarnings++;
                    }
                    catch (Exception ex) when (!(ex is OutOfMemoryException))
                    {
                        failed++;
                        if (failures.Count < 20) failures.Add($"0x{work[i].Value:X}: {ex.Message}");
                    }
                }
                Console.Write("\r" + new string(' ', 40) + "\r");

                Console.WriteLine();
                foreach (var f in chosen)
                    Console.WriteLine($"Folder: {f.Path}  ({f.Listed.Count:N0} listed{(f.Extra.Count > 0 ? $", {f.Extra.Count:N0} more found outside the list" : "")})");
                Console.WriteLine($"Messages that can be exported: {ok:N0}");
                if (withWarnings > 0) Console.WriteLine($"   of which {withWarnings:N0} have a damaged part (some recipients or attachments may be missing)");
                if (failed > 0) Console.WriteLine($"Messages that cannot be read: {failed:N0}");
                if (missing > 0) Console.WriteLine($"Listed messages whose data is gone from the file: {missing:N0}");
                Console.WriteLine($"Size of the .msg files: {Size(total)} ({total:N0} bytes); about {Size(onDisk)} on disk");
                foreach (var f in failures) Console.WriteLine("   " + f);

                if (ok == 0) { Warn("Nothing to export."); return 0; }
                if (outDir == null)
                    outDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory), Safe($"{Path.GetFileNameWithoutExtension(file)} - {(folderArg == "*" ? "All folders" : folderArg.Replace('/', ' '))} export {DateTime.Now:yyyy-MM-dd HHmm}"));
                outDir = Path.GetFullPath(outDir);
                var drive = new DriveInfo(Path.GetPathRoot(outDir));
                Console.WriteLine($"Destination: {outDir}");
                Console.WriteLine($"Free space there: {Size(drive.AvailableFreeSpace)}");
                if (drive.AvailableFreeSpace < onDisk + 50L * 1024 * 1024) { Warn("There is not enough free space for the export."); return 4; }
                Console.WriteLine($"Analysis took {sw.Elapsed.TotalSeconds:N0} s.");
                Console.WriteLine();
                if (!yes)
                {
                    Console.Write($"Export {ok:N0} messages ({Size(total)}) now? [Y/N] ");
                    var ans = (Console.ReadLine() ?? "").Trim();
                    if (!ans.Equals("y", StringComparison.OrdinalIgnoreCase) && !ans.Equals("yes", StringComparison.OrdinalIgnoreCase)) { Console.WriteLine("Nothing was exported."); return 0; }
                }

                // ---------- export ----------
                Directory.CreateDirectory(outDir);
                var log = new StringBuilder();
                log.AppendLine($"OstMsgExport {Version}  {DateTime.Now:yyyy-MM-dd HH:mm}");
                log.AppendLine($"Source: {file}");
                int written = 0, errors = 0; long writtenBytes = 0;
                var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                for (int i = 0; i < work.Count; i++)
                {
                    if (i % 25 == 0) Console.Write($"\r   writing {i:N0} of {work.Count:N0}");
                    var f = work[i].Key; uint nid = work[i].Value;
                    if (!ost.Nodes.TryGetValue(nid, out var e)) continue;
                    try
                    {
                        var m = MessageReader.Read(new Node(ost, e));
                        var bytes = MsgBuilder.Build(m);
                        string dir = chosen.Count > 1 ? Path.Combine(outDir, Safe(f.Path.Replace('/', '_'))) : outDir;
                        if (f.Extra.Contains(nid)) dir = Path.Combine(dir, "Found outside the folder list");
                        Directory.CreateDirectory(dir);
                        var when = m.GetTime(0x0E06) ?? m.GetTime(0x0039) ?? m.GetTime(0x3007);
                        string name = $"{(when.HasValue ? when.Value.ToString("yyyy-MM-dd HHmm") : "no date")} {m.GetString(0x0C1A) ?? m.GetString(0x0042) ?? ""} - {m.GetString(0x0037) ?? "(no subject)"}";
                        name = Safe(name);
                        if (name.Length > 120) name = name.Substring(0, 120).TrimEnd();
                        string path = Path.Combine(dir, name + ".msg");
                        for (int k = 2; used.Contains(path) || File.Exists(path); k++) path = Path.Combine(dir, $"{name} ({k}).msg");
                        used.Add(path);
                        File.WriteAllBytes(path, bytes);
                        if (when.HasValue) { try { File.SetCreationTime(path, when.Value); File.SetLastWriteTime(path, when.Value); } catch { } }
                        written++; writtenBytes += bytes.Length;
                        foreach (var w in m.Warnings) log.AppendLine($"WARNING {Path.GetFileName(path)}: {w}");
                    }
                    catch (Exception ex) when (!(ex is OutOfMemoryException))
                    {
                        errors++;
                        log.AppendLine($"FAILED message 0x{nid:X}: {ex.Message}");
                    }
                }
                Console.Write("\r" + new string(' ', 40) + "\r");
                log.AppendLine($"Written: {written} files, {writtenBytes:N0} bytes. Failed: {errors}.");
                File.WriteAllText(Path.Combine(outDir, "export-log.txt"), log.ToString(), Encoding.UTF8);
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Exported {written:N0} messages ({Size(writtenBytes)}) to {outDir}");
                Console.ResetColor();
                if (errors > 0) Warn($"{errors:N0} messages could not be written; see export-log.txt.");
                return errors > 0 ? 5 : 0;
            }
        }

        static List<FolderInfo> LoadFolders(OstFile ost)
        {
            var folders = new Dictionary<uint, FolderInfo>();
            foreach (var e in ost.Nodes.Values)
            {
                if (e.Type != 0x02) continue;
                var f = new FolderInfo { Nid = e.Nid, Name = "(unnamed)" };
                try
                {
                    foreach (var p in PropertyContext.Read(new Node(ost, e)))
                    {
                        if (p.Tag == 0x3001001F) f.Name = Encoding.Unicode.GetString(p.Value).TrimEnd('\0');
                        else if (p.Tag == 0x3001001E) f.Name = Encoding.Default.GetString(p.Value).TrimEnd('\0');
                        else if (p.Tag == 0x36020003) f.ContentCount = BitConverter.ToInt32(p.Value, 0);
                    }
                }
                catch (Exception ex) when (ex is InvalidDataException || ex is ArgumentException || ex is IndexOutOfRangeException) { f.Name = $"(damaged folder 0x{e.Nid:X})"; }
                uint contents = (e.Nid & ~0x1Fu) | 0x0E;
                if (ost.Nodes.TryGetValue(contents, out var ce))
                {
                    try { f.Listed.AddRange(new TableContext(new Node(ost, ce)).RowIds); }
                    catch (Exception ex) when (ex is InvalidDataException || ex is ArgumentException || ex is IndexOutOfRangeException) { }
                }
                folders[e.Nid] = f;
            }
            foreach (var f in folders.Values)
            {
                var listed = new HashSet<uint>(f.Listed);
                foreach (var e in ost.Nodes.Values)
                    if (e.Type == 0x04 && e.NidParent == f.Nid && !listed.Contains(e.Nid)) f.Extra.Add(e.Nid);
            }
            // Paths, from the parent links.
            foreach (var f in folders.Values)
            {
                var parts = new List<string>();
                uint cur = f.Nid;
                for (int guard = 0; guard < 32; guard++)
                {
                    if (!folders.TryGetValue(cur, out var c)) break;
                    if (cur == 0x122) break; // the root folder itself has no useful name
                    parts.Insert(0, c.Name);
                    if (!ost.Nodes.TryGetValue(cur, out var ne) || ne.NidParent == cur) break;
                    cur = ne.NidParent;
                }
                f.Path = parts.Count > 0 ? string.Join("/", parts) : "(root)";
            }
            return folders.Values.ToList();
        }

        static List<FolderInfo> SelectFolders(List<FolderInfo> folders, string arg)
        {
            if (arg == "*") return folders.Where(f => f.Listed.Count + f.Extra.Count > 0).OrderBy(f => f.Path).ToList();
            var byPath = folders.Where(f => f.Path.Equals(arg, StringComparison.OrdinalIgnoreCase) || f.Path.EndsWith("/" + arg, StringComparison.OrdinalIgnoreCase)).ToList();
            if (byPath.Count > 0) return byPath;
            return folders.Where(f => f.Name.Equals(arg, StringComparison.OrdinalIgnoreCase)).ToList();
        }

        static string Safe(string s)
        {
            var bad = Path.GetInvalidFileNameChars();
            var sb = new StringBuilder();
            foreach (var ch in s) sb.Append(Array.IndexOf(bad, ch) >= 0 || ch < 32 ? '_' : ch);
            return sb.ToString().Trim().TrimEnd('.');
        }

        static string Size(long b) => b >= 1L << 30 ? $"{b / (double)(1L << 30):N2} GB" : b >= 1L << 20 ? $"{b / (double)(1L << 20):N1} MB" : $"{b / 1024.0:N0} KB";

        static void Warn(string s) { Console.ForegroundColor = ConsoleColor.Yellow; Console.WriteLine(s); Console.ResetColor(); }
    }
}
