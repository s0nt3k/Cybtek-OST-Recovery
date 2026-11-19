// Minimal compound file (OLE structured storage, [MS-CFB] version 3) writer, as used by Outlook .msg files.
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace OstMsgExport
{
    sealed class CfbStorage
    {
        public readonly string Name;
        public readonly List<CfbStorage> Storages = new List<CfbStorage>();
        public readonly List<KeyValuePair<string, byte[]>> Streams = new List<KeyValuePair<string, byte[]>>();
        public CfbStorage(string name) { Name = name; }
        public CfbStorage AddStorage(string name) { var s = new CfbStorage(name); Storages.Add(s); return s; }
        public void AddStream(string name, byte[] data) => Streams.Add(new KeyValuePair<string, byte[]>(name, data ?? new byte[0]));
    }

    static class CfbWriter
    {
        const uint FREESECT = 0xFFFFFFFF, ENDOFCHAIN = 0xFFFFFFFE, FATSECT = 0xFFFFFFFD, DIFSECT = 0xFFFFFFFC, NOSTREAM = 0xFFFFFFFF;
        const int Sector = 512, MiniSector = 64, MiniCutoff = 4096;

        sealed class Entry
        {
            public string Name; public byte Type; public byte[] Data; public uint Left = NOSTREAM, Right = NOSTREAM, Child = NOSTREAM;
            public uint Start = ENDOFCHAIN; public long Size;
        }

        static int Compare(string a, string b)
        {
            if (a.Length != b.Length) return a.Length.CompareTo(b.Length);
            return string.CompareOrdinal(a.ToUpperInvariant(), b.ToUpperInvariant());
        }

        public static byte[] Write(CfbStorage root)
        {
            var entries = new List<Entry> { new Entry { Name = "Root Entry", Type = 5 } };
            AddChildren(root, 0, entries);

            // Small streams go into the mini stream.
            var mini = new MemoryStream();
            foreach (var e in entries.Where(x => x.Type == 2 && x.Data.Length > 0 && x.Data.Length < MiniCutoff))
            {
                e.Start = (uint)(mini.Length / MiniSector);
                e.Size = e.Data.Length;
                mini.Write(e.Data, 0, e.Data.Length);
                int pad = (int)(-mini.Length & (MiniSector - 1)); if (pad > 0) mini.Write(new byte[pad], 0, pad);
            }
            foreach (var e in entries.Where(x => x.Type == 2 && x.Data.Length == 0)) { e.Start = ENDOFCHAIN; e.Size = 0; }
            var miniFat = new List<uint>();
            foreach (var e in entries.Where(x => x.Type == 2 && x.Data.Length > 0 && x.Data.Length < MiniCutoff))
            {
                int n = (e.Data.Length + MiniSector - 1) / MiniSector;
                for (int i = 0; i < n; i++) miniFat.Add(i == n - 1 ? ENDOFCHAIN : e.Start + (uint)i + 1);
            }

            // Regular sectors: big streams, mini stream, mini FAT, directory, then FAT and DIFAT.
            var sectors = new List<byte[]>();
            var fat = new List<uint>();
            uint AddChain(byte[] data)
            {
                if (data.Length == 0) return ENDOFCHAIN;
                uint first = (uint)sectors.Count;
                int n = (data.Length + Sector - 1) / Sector;
                for (int i = 0; i < n; i++)
                {
                    var s = new byte[Sector];
                    Buffer.BlockCopy(data, i * Sector, s, 0, Math.Min(Sector, data.Length - i * Sector));
                    sectors.Add(s);
                    fat.Add(i == n - 1 ? ENDOFCHAIN : first + (uint)i + 1);
                }
                return first;
            }
            foreach (var e in entries.Where(x => x.Type == 2 && x.Data.Length >= MiniCutoff)) { e.Start = AddChain(e.Data); e.Size = e.Data.Length; }
            entries[0].Start = mini.Length > 0 ? AddChain(mini.ToArray()) : ENDOFCHAIN;
            entries[0].Size = mini.Length;
            byte[] miniFatBytes = new byte[miniFat.Count * 4];
            for (int i = 0; i < miniFat.Count; i++) BitConverter.GetBytes(miniFat[i]).CopyTo(miniFatBytes, i * 4);
            if (miniFatBytes.Length % Sector != 0) { int old = miniFatBytes.Length; Array.Resize(ref miniFatBytes, (old / Sector + 1) * Sector); for (int i = old; i < miniFatBytes.Length; i += 4) BitConverter.GetBytes(FREESECT).CopyTo(miniFatBytes, i); }
            uint miniFatStart = AddChain(miniFatBytes);
            int miniFatSectors = miniFatBytes.Length / Sector;

            var dir = new byte[((entries.Count + 3) / 4) * 4 * 128];
            for (int i = 0; i < dir.Length / 128; i++)
            {
                int o = i * 128;
                if (i >= entries.Count) { BitConverter.GetBytes(NOSTREAM).CopyTo(dir, o + 68); BitConverter.GetBytes(NOSTREAM).CopyTo(dir, o + 72); BitConverter.GetBytes(NOSTREAM).CopyTo(dir, o + 76); continue; }
                var e = entries[i];
                var nb = Encoding.Unicode.GetBytes(e.Name);
                Buffer.BlockCopy(nb, 0, dir, o, Math.Min(nb.Length, 62));
                BitConverter.GetBytes((ushort)(Math.Min(nb.Length, 62) + 2)).CopyTo(dir, o + 64);
                dir[o + 66] = e.Type; dir[o + 67] = 1; // black
                BitConverter.GetBytes(e.Left).CopyTo(dir, o + 68);
                BitConverter.GetBytes(e.Right).CopyTo(dir, o + 72);
                BitConverter.GetBytes(e.Child).CopyTo(dir, o + 76);
                BitConverter.GetBytes(e.Type == 1 ? 0u : e.Start).CopyTo(dir, o + 116);
                BitConverter.GetBytes(e.Type == 1 ? 0UL : (ulong)e.Size).CopyTo(dir, o + 120);
            }
            uint dirStart = AddChain(dir);

            // FAT sectors (and DIFAT sectors beyond the 109 entries in the header).
            int nData = sectors.Count, nFat = 0, nDifat = 0;
            while (true)
            {
                int needDifat = nFat > 109 ? (nFat - 109 + 126) / 127 : 0;
                int needFat = (nData + nFat + needDifat + 127) / 128;
                if (needFat == nFat && needDifat == nDifat) break;
                nFat = needFat; nDifat = needDifat;
            }
            uint fatStart = (uint)nData, difatStart = (uint)(nData + nFat);
            for (int i = 0; i < nFat; i++) fat.Add(FATSECT);
            for (int i = 0; i < nDifat; i++) fat.Add(DIFSECT);
            while (fat.Count < nFat * 128) fat.Add(FREESECT);
            for (int i = 0; i < nFat; i++)
            {
                var s = new byte[Sector];
                for (int j = 0; j < 128; j++) BitConverter.GetBytes(fat[i * 128 + j]).CopyTo(s, j * 4);
                sectors.Add(s);
            }
            for (int d = 0; d < nDifat; d++)
            {
                var s = new byte[Sector];
                for (int j = 0; j < 127; j++)
                {
                    int fi = 109 + d * 127 + j;
                    BitConverter.GetBytes(fi < nFat ? fatStart + (uint)fi : FREESECT).CopyTo(s, j * 4);
                }
                BitConverter.GetBytes(d == nDifat - 1 ? ENDOFCHAIN : difatStart + (uint)d + 1).CopyTo(s, 508);
                sectors.Add(s);
            }

            var hdr = new byte[Sector];
            new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }.CopyTo(hdr, 0);
            BitConverter.GetBytes((ushort)0x3E).CopyTo(hdr, 24);
            BitConverter.GetBytes((ushort)3).CopyTo(hdr, 26);
            BitConverter.GetBytes((ushort)0xFFFE).CopyTo(hdr, 28);
            BitConverter.GetBytes((ushort)9).CopyTo(hdr, 30);
            BitConverter.GetBytes((ushort)6).CopyTo(hdr, 32);
            BitConverter.GetBytes((uint)nFat).CopyTo(hdr, 44);
            BitConverter.GetBytes(dirStart).CopyTo(hdr, 48);
            BitConverter.GetBytes((uint)MiniCutoff).CopyTo(hdr, 56);
            BitConverter.GetBytes(miniFatSectors > 0 ? miniFatStart : ENDOFCHAIN).CopyTo(hdr, 60);
            BitConverter.GetBytes((uint)miniFatSectors).CopyTo(hdr, 64);
            BitConverter.GetBytes(nDifat > 0 ? difatStart : ENDOFCHAIN).CopyTo(hdr, 68);
            BitConverter.GetBytes((uint)nDifat).CopyTo(hdr, 72);
            for (int i = 0; i < 109; i++) BitConverter.GetBytes(i < nFat ? fatStart + (uint)i : FREESECT).CopyTo(hdr, 76 + i * 4);

            var output = new byte[Sector * (1 + sectors.Count)];
            Buffer.BlockCopy(hdr, 0, output, 0, Sector);
            for (int i = 0; i < sectors.Count; i++) Buffer.BlockCopy(sectors[i], 0, output, Sector * (i + 1), Sector);
            return output;
        }

        static void AddChildren(CfbStorage s, int parentIndex, List<Entry> entries)
        {
            var kids = new List<KeyValuePair<Entry, CfbStorage>>();
            foreach (var st in s.Storages) kids.Add(new KeyValuePair<Entry, CfbStorage>(new Entry { Name = st.Name, Type = 1 }, st));
            foreach (var kv in s.Streams) kids.Add(new KeyValuePair<Entry, CfbStorage>(new Entry { Name = kv.Key, Type = 2, Data = kv.Value }, null));
            if (kids.Count == 0) return;
            kids.Sort((a, b) => Compare(a.Key.Name, b.Key.Name));
            int baseIndex = entries.Count;
            foreach (var k in kids) entries.Add(k.Key);
            entries[parentIndex].Child = (uint)(baseIndex + Balance(kids, 0, kids.Count - 1, baseIndex, entries));
            for (int i = 0; i < kids.Count; i++) if (kids[i].Value != null) AddChildren(kids[i].Value, baseIndex + i, entries);
        }

        // Builds a balanced binary search tree over kids[lo..hi]; returns the index (relative to baseIndex) of its root.
        static int Balance(List<KeyValuePair<Entry, CfbStorage>> kids, int lo, int hi, int baseIndex, List<Entry> entries)
        {
            int mid = (lo + hi) / 2;
            var e = entries[baseIndex + mid];
            if (lo <= mid - 1) e.Left = (uint)(baseIndex + Balance(kids, lo, mid - 1, baseIndex, entries));
            if (mid + 1 <= hi) e.Right = (uint)(baseIndex + Balance(kids, mid + 1, hi, baseIndex, entries));
            return mid;
        }
    }
}
