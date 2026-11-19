// Builds Outlook .msg files ([MS-OXMSG]) from messages read out of the OST.
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OstMsgExport
{
    sealed class MessageData
    {
        public List<Prop> Props = new List<Prop>();
        public List<List<Prop>> Recipients = new List<List<Prop>>();
        public List<AttachmentData> Attachments = new List<AttachmentData>();
        public List<string> Warnings = new List<string>();

        public Prop Get(uint tag) { foreach (var p in Props) if (p.Tag == tag) return p; return null; }
        public string GetString(ushort id)
        {
            foreach (var p in Props)
                if (p.Id == id && p.Value != null)
                {
                    if (p.Type == PropTypes.Unicode) return Encoding.Unicode.GetString(p.Value).TrimEnd('\0');
                    if (p.Type == PropTypes.String8) return Encoding.Default.GetString(p.Value).TrimEnd('\0');
                }
            return null;
        }
        public DateTime? GetTime(ushort id)
        {
            var p = Get(((uint)id << 16) | PropTypes.SysTime);
            if (p?.Value == null || p.Value.Length < 8) return null;
            long ft = BitConverter.ToInt64(p.Value, 0);
            try { return ft > 0 ? DateTime.FromFileTimeUtc(ft).ToLocalTime() : (DateTime?)null; } catch { return null; }
        }
    }

    sealed class AttachmentData
    {
        public List<Prop> Props = new List<Prop>();
        public MessageData Embedded;
    }

    static class MessageReader
    {
        const uint NidRecipients = 0x692, NidAttachments = 0x671;
        const uint TagAttachDataObject = 0x3701000D, TagAttachMethod = 0x37050003;

        public static MessageData Read(Node node, int depth = 0)
        {
            if (depth > 5) throw new InvalidDataException("Attachments are nested too deeply.");
            var m = new MessageData();
            foreach (var p in PropertyContext.Read(node)) if (Keep(p)) m.Props.Add(p);
            CleanSubject(m.Props);

            var rn = node.Subnode(NidRecipients);
            if (rn != null)
            {
                try { foreach (var row in new TableContext(rn).Rows()) m.Recipients.Add(row.FindAll(Keep)); }
                catch (Exception e) when (e is InvalidDataException || e is ArgumentException || e is IndexOutOfRangeException) { m.Warnings.Add("recipients unreadable: " + e.Message); }
            }

            var an = node.Subnode(NidAttachments);
            if (an != null)
            {
                List<uint> ids;
                try { ids = new TableContext(an).RowIds; }
                catch (Exception e) when (e is InvalidDataException || e is ArgumentException || e is IndexOutOfRangeException) { ids = new List<uint>(); m.Warnings.Add("attachment list unreadable: " + e.Message); }
                foreach (var aid in ids)
                {
                    try
                    {
                        var anode = node.Subnode(aid);
                        if (anode == null) { m.Warnings.Add($"attachment 0x{aid:X} missing"); continue; }
                        var a = new AttachmentData();
                        Prop obj = null; int method = 0;
                        foreach (var p in PropertyContext.Read(anode))
                        {
                            if (p.Tag == TagAttachDataObject) { obj = p; continue; }
                            if (p.Tag == TagAttachMethod && p.Value != null) method = BitConverter.ToInt32(p.Value, 0);
                            if (Keep(p)) a.Props.Add(p);
                        }
                        if (obj != null && method == 5)
                        {
                            var en = anode.Subnode(obj.ObjectNid);
                            if (en != null) a.Embedded = Read(en, depth + 1);
                            else m.Warnings.Add("attached message missing");
                        }
                        else if (obj != null)
                        {
                            // OLE object data: stored as binary so it is not lost.
                            var on = anode.Subnode(obj.ObjectNid);
                            if (on != null) a.Props.Add(new Prop { Tag = 0x37010102, Value = on.Data() });
                        }
                        m.Attachments.Add(a);
                    }
                    catch (Exception e) when (e is InvalidDataException || e is ArgumentException || e is IndexOutOfRangeException) { m.Warnings.Add($"attachment 0x{aid:X} unreadable: {e.Message}"); }
                }
            }
            return m;
        }

        // Named properties (id >= 0x8000) depend on the source file's name map and are not exported.
        static bool Keep(Prop p) => p.Id < 0x8000 && (p.Value != null || p.Type == PropTypes.Object);

        /// <summary>Subjects in the file may start with U+0001 and a prefix length; .msg files hold the plain subject.</summary>
        public static void CleanSubject(List<Prop> props)
        {
            foreach (var p in props)
                if (p.Tag == 0x0037001F && p.Value != null && p.Value.Length >= 4 && p.Value[0] == 1 && p.Value[1] == 0)
                {
                    var v = new byte[p.Value.Length - 4]; Buffer.BlockCopy(p.Value, 4, v, 0, v.Length); p.Value = v;
                }
                else if (p.Tag == 0x0037001E && p.Value != null && p.Value.Length >= 2 && p.Value[0] == 1)
                {
                    var v = new byte[p.Value.Length - 2]; Buffer.BlockCopy(p.Value, 2, v, 0, v.Length); p.Value = v;
                }
        }
    }

    static class MsgBuilder
    {
        public static byte[] Build(MessageData m)
        {
            var root = new CfbStorage("Root Entry");
            Fill(root, m, true);
            var names = root.AddStorage("__nameid_version1.0");
            names.AddStream("__substg1.0_00020102", new byte[0]);
            names.AddStream("__substg1.0_00030102", new byte[0]);
            names.AddStream("__substg1.0_00040102", new byte[0]);
            return CfbWriter.Write(root);
        }

        static void Fill(CfbStorage st, MessageData m, bool topLevel)
        {
            // Outlook reads string properties as Unicode only when the store support mask says so.
            if (m.Get(0x340D0003) == null) m.Props.Add(new Prop { Tag = 0x340D0003, Value = BitConverter.GetBytes(0x00040000) });
            var header = new byte[topLevel ? 32 : 24];
            BitConverter.GetBytes(m.Recipients.Count).CopyTo(header, 8);
            BitConverter.GetBytes(m.Attachments.Count).CopyTo(header, 12);
            BitConverter.GetBytes(m.Recipients.Count).CopyTo(header, 16);
            BitConverter.GetBytes(m.Attachments.Count).CopyTo(header, 20);
            WriteProps(st, header, m.Props);

            for (int i = 0; i < m.Recipients.Count; i++)
            {
                var rs = st.AddStorage($"__recip_version1.0_#{i:X8}");
                var props = new List<Prop>(m.Recipients[i]);
                if (!props.Exists(p => p.Tag == 0x30000003)) props.Add(new Prop { Tag = 0x30000003, Value = BitConverter.GetBytes(i) });
                WriteProps(rs, new byte[8], props);
            }
            for (int i = 0; i < m.Attachments.Count; i++)
            {
                var a = m.Attachments[i];
                var ast = st.AddStorage($"__attach_version1.0_#{i:X8}");
                var props = new List<Prop>(a.Props);
                if (!props.Exists(p => p.Tag == 0x0E210003)) props.Add(new Prop { Tag = 0x0E210003, Value = BitConverter.GetBytes(i) });
                if (a.Embedded != null)
                {
                    props.Add(new Prop { Tag = 0x3701000D });
                    Fill(ast.AddStorage("__substg1.0_3701000D"), a.Embedded, false);
                }
                WriteProps(ast, new byte[8], props);
            }
        }

        static void WriteProps(CfbStorage st, byte[] header, List<Prop> props)
        {
            var ms = new MemoryStream();
            ms.Write(header, 0, header.Length);
            var seen = new HashSet<uint>();
            foreach (var p in props)
            {
                if (!seen.Add(p.Tag)) continue;
                var entry = new byte[16];
                BitConverter.GetBytes(p.Tag).CopyTo(entry, 0);
                BitConverter.GetBytes(6u).CopyTo(entry, 4); // readable | writable
                int fixedSize = PropTypes.FixedSize(p.Type);
                if (p.Type == PropTypes.Object)
                {
                    BitConverter.GetBytes(0xFFFFFFFFu).CopyTo(entry, 8);
                    BitConverter.GetBytes(4u).CopyTo(entry, 12);
                }
                else if (fixedSize > 0)
                {
                    var v = p.Value;
                    if (p.Type == PropTypes.Boolean) v = new byte[] { (byte)(v.Length > 0 && v[0] != 0 ? 1 : 0), 0 };
                    Buffer.BlockCopy(v, 0, entry, 8, Math.Min(8, v.Length));
                }
                else
                {
                    var data = p.Value ?? new byte[0];
                    // Strings are stored without the terminator; the size field counts it.
                    if (p.Type == PropTypes.Unicode) { data = TrimTerminator(data, 2); BitConverter.GetBytes((uint)data.Length + 2).CopyTo(entry, 8); }
                    else if (p.Type == PropTypes.String8) { data = TrimTerminator(data, 1); BitConverter.GetBytes((uint)data.Length + 1).CopyTo(entry, 8); }
                    else BitConverter.GetBytes((uint)data.Length).CopyTo(entry, 8);
                    st.AddStream($"__substg1.0_{p.Tag:X8}", data);
                }
                ms.Write(entry, 0, 16);
            }
            st.AddStream("__properties_version1.0", ms.ToArray());
        }

        static byte[] TrimTerminator(byte[] d, int unit)
        {
            int n = d.Length - d.Length % unit;
            while (n >= unit)
            {
                bool zero = true;
                for (int i = n - unit; i < n; i++) if (d[i] != 0) { zero = false; break; }
                if (!zero) break;
                n -= unit;
            }
            if (n == d.Length) return d;
            var r = new byte[n]; Buffer.BlockCopy(d, 0, r, 0, n); return r;
        }
    }
}
