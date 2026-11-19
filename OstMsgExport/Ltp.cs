// Lists, tables and properties layer ([MS-PST] 2.3): heap-on-node, BTH, property contexts and table contexts.
using System;
using System.Collections.Generic;
using System.IO;

namespace OstMsgExport
{
    /// <summary>One MAPI property value. Fixed values are little-endian bytes; variable values are the raw bytes.</summary>
    sealed class Prop
    {
        public uint Tag;                 // (id << 16) | type
        public ushort Id => (ushort)(Tag >> 16);
        public ushort Type => (ushort)(Tag & 0xFFFF);
        public byte[] Value;             // null for an object reference
        public uint ObjectNid;           // PtypObject: subnode holding the object
        public uint ObjectSize;
    }

    static class PropTypes
    {
        public const ushort Int16 = 0x0002, Int32 = 0x0003, Float = 0x0004, Double = 0x0005, Currency = 0x0006, AppTime = 0x0007,
            Error = 0x000A, Boolean = 0x000B, Object = 0x000D, Int64 = 0x0014, String8 = 0x001E, Unicode = 0x001F,
            SysTime = 0x0040, Guid = 0x0048, ServerId = 0x00FB, Restriction = 0x00FD, RuleAction = 0x00FE, Binary = 0x0102;

        /// <summary>Size of a fixed-length type, or 0 for variable-length / unsupported.</summary>
        public static int FixedSize(ushort t)
        {
            switch (t)
            {
                case Int16: return 2;
                case Int32: case Float: case Error: return 4;
                case Boolean: return 1;
                case Double: case Currency: case AppTime: case Int64: case SysTime: return 8;
                default: return 0;
            }
        }

        public static bool IsVariable(ushort t) => t == String8 || t == Unicode || t == Binary || t == Guid || t == ServerId;
    }

    sealed class Heap
    {
        readonly Node node;
        readonly List<byte[]> blocks;
        public readonly byte ClientSig;
        public readonly uint UserRoot;

        public Heap(Node node)
        {
            this.node = node;
            blocks = node.Blocks;
            if (blocks.Count == 0) throw new InvalidDataException("Empty node.");
            var b0 = blocks[0];
            if (b0.Length < 12 || b0[2] != 0xEC) throw new InvalidDataException($"Node 0x{node.Nid:X} is not a heap (signature 0x{(b0.Length > 2 ? b0[2] : 0):X2}).");
            ClientSig = b0[3];
            UserRoot = BitConverter.ToUInt32(b0, 4);
        }

        public byte[] Item(uint hid)
        {
            if (hid == 0) return null;
            // Unicode files: 11-bit item index, block index from bit 16. 4K files (up to 64 KB blocks): 14-bit index, block from bit 19.
            int blockIndex, index;
            if (node.Is4K) { blockIndex = (int)(hid >> 19); index = (int)((hid >> 5) & 0x3FFF); }
            else { blockIndex = (int)(hid >> 16); index = (int)((hid >> 5) & 0x7FF); }
            if (blockIndex >= blocks.Count || index == 0) throw new InvalidDataException($"Heap item 0x{hid:X} is out of range.");
            var b = blocks[blockIndex];
            int map = BitConverter.ToUInt16(b, 0);
            int cAlloc = BitConverter.ToUInt16(b, map);
            if (index > cAlloc) throw new InvalidDataException($"Heap item 0x{hid:X} is out of range.");
            int start = BitConverter.ToUInt16(b, map + 4 + (index - 1) * 2), end = BitConverter.ToUInt16(b, map + 4 + index * 2);
            if (end < start || end > b.Length) throw new InvalidDataException($"Heap item 0x{hid:X} is damaged.");
            var r = new byte[end - start];
            Buffer.BlockCopy(b, start, r, 0, r.Length);
            return r;
        }

        /// <summary>A heap item, or the data of a subnode when the value is a NID.</summary>
        public byte[] Resolve(uint hnid)
        {
            if (hnid == 0) return new byte[0];
            if ((hnid & 0x1F) == 0) return Item(hnid);
            var sn = node.Subnode(hnid);
            if (sn == null) throw new InvalidDataException($"Subnode 0x{hnid:X} is missing.");
            return sn.Data();
        }

        public Node Node => node;

        /// <summary>All leaf records (key, data) of a BTH whose header is the heap item at hidHeader.</summary>
        public List<KeyValuePair<byte[], byte[]>> ReadBth(uint hidHeader)
        {
            var list = new List<KeyValuePair<byte[], byte[]>>();
            var h = Item(hidHeader);
            if (h == null || h.Length < 8 || h[0] != 0xB5) throw new InvalidDataException("Damaged BTH header.");
            int cbKey = h[1], cbEnt = h[2], levels = h[3];
            uint root = BitConverter.ToUInt32(h, 4);
            if (root != 0) AddBth(root, cbKey, cbEnt, levels, list, 0);
            return list;
        }

        void AddBth(uint hid, int cbKey, int cbEnt, int level, List<KeyValuePair<byte[], byte[]>> list, int depth)
        {
            if (depth > 8) throw new InvalidDataException("BTH too deep.");
            var d = Item(hid);
            int rec = level > 0 ? cbKey + 4 : cbKey + cbEnt;
            for (int o = 0; o + rec <= d.Length; o += rec)
            {
                if (level > 0) { AddBth(BitConverter.ToUInt32(d, o + cbKey), cbKey, cbEnt, level - 1, list, depth + 1); continue; }
                var k = new byte[cbKey]; var v = new byte[cbEnt];
                Buffer.BlockCopy(d, o, k, 0, cbKey); Buffer.BlockCopy(d, o + cbKey, v, 0, cbEnt);
                list.Add(new KeyValuePair<byte[], byte[]>(k, v));
            }
        }
    }

    static class PropertyContext
    {
        public static List<Prop> Read(Node node)
        {
            var heap = new Heap(node);
            if (heap.ClientSig != 0xBC) throw new InvalidDataException($"Node 0x{node.Nid:X} is not a property context.");
            var list = new List<Prop>();
            foreach (var rec in heap.ReadBth(heap.UserRoot))
            {
                ushort id = BitConverter.ToUInt16(rec.Key, 0);
                ushort type = BitConverter.ToUInt16(rec.Value, 0);
                uint hnid = BitConverter.ToUInt32(rec.Value, 2);
                var p = Decode(heap, id, type, hnid, false);
                if (p != null) list.Add(p);
            }
            return list;
        }

        /// <summary>Decodes one value. inRow: the value came from a table row, where 8-byte values are stored inline.</summary>
        public static Prop Decode(Heap heap, ushort id, ushort type, uint hnidOrValue, bool inRow, byte[] inline8 = null)
        {
            var p = new Prop { Tag = ((uint)id << 16) | type };
            int fixedSize = PropTypes.FixedSize(type);
            if (fixedSize > 0 && fixedSize <= 4)
            {
                p.Value = BitConverter.GetBytes(hnidOrValue);
                Array.Resize(ref p.Value, fixedSize);
                return p;
            }
            if (fixedSize == 8)
            {
                p.Value = inRow ? inline8 : heap.Resolve(hnidOrValue);
                if (p.Value == null || p.Value.Length < 8) return null;
                if (p.Value.Length > 8) Array.Resize(ref p.Value, 8);
                return p;
            }
            if (type == PropTypes.Object)
            {
                var d = heap.Resolve(hnidOrValue);
                if (d == null || d.Length < 8) return null;
                p.ObjectNid = BitConverter.ToUInt32(d, 0);
                p.ObjectSize = BitConverter.ToUInt32(d, 4);
                return p;
            }
            if (PropTypes.IsVariable(type))
            {
                p.Value = heap.Resolve(hnidOrValue) ?? new byte[0];
                return p;
            }
            return null; // multi-valued and other types are not exported
        }
    }

    /// <summary>A table context: column descriptions and rows.</summary>
    sealed class TableContext
    {
        sealed class Col { public uint Tag; public int Ib; public int Cb; public int Bit; }
        readonly Heap heap;
        readonly List<Col> cols = new List<Col>();
        readonly int rowWidth, cebOffset;
        readonly List<byte[]> rowBlocks = new List<byte[]>();
        public readonly List<uint> RowIds = new List<uint>();
        readonly List<int> rowIndexes = new List<int>();

        public TableContext(Node node)
        {
            heap = new Heap(node);
            if (heap.ClientSig != 0x7C) throw new InvalidDataException($"Node 0x{node.Nid:X} is not a table.");
            var h = heap.Item(heap.UserRoot);
            if (h[0] != 0x7C) throw new InvalidDataException("Damaged table header.");
            int cCols = h[1];
            cebOffset = BitConverter.ToUInt16(h, 6);   // TCI_1b: end of 1-byte columns = start of the existence bitmap
            rowWidth = BitConverter.ToUInt16(h, 8);    // TCI_bm: total row size
            uint hidRowIndex = BitConverter.ToUInt32(h, 10), hnidRows = BitConverter.ToUInt32(h, 14);
            for (int i = 0; i < cCols; i++)
            {
                int o = 22 + i * 8;
                cols.Add(new Col { Tag = BitConverter.ToUInt32(h, o), Ib = BitConverter.ToUInt16(h, o + 4), Cb = h[o + 6], Bit = h[o + 7] });
            }
            foreach (var rec in heap.ReadBth(hidRowIndex))
            {
                RowIds.Add(BitConverter.ToUInt32(rec.Key, 0));
                rowIndexes.Add(rec.Value.Length >= 4 ? BitConverter.ToInt32(rec.Value, 0) : BitConverter.ToUInt16(rec.Value, 0));
            }
            if (hnidRows == 0 || rowWidth == 0) return;
            if ((hnidRows & 0x1F) == 0) rowBlocks.Add(heap.Item(hnidRows));
            else
            {
                var sn = node.Subnode(hnidRows);
                if (sn != null) rowBlocks.AddRange(sn.Blocks);
            }
        }

        byte[] Row(int rowIndex)
        {
            if (rowBlocks.Count == 0) return null;
            int perBlock = rowBlocks.Count == 1 ? int.MaxValue : rowBlocks[0].Length / rowWidth;
            int bi = rowBlocks.Count == 1 ? 0 : rowIndex / perBlock, ri = rowBlocks.Count == 1 ? rowIndex : rowIndex % perBlock;
            if (bi >= rowBlocks.Count) return null;
            int o = ri * rowWidth;
            if (o + rowWidth > rowBlocks[bi].Length) return null;
            var r = new byte[rowWidth];
            Buffer.BlockCopy(rowBlocks[bi], o, r, 0, rowWidth);
            return r;
        }

        /// <summary>The properties of each row, in row-index order.</summary>
        public List<List<Prop>> Rows()
        {
            var all = new List<List<Prop>>();
            for (int i = 0; i < RowIds.Count; i++)
            {
                var row = Row(rowIndexes[i]);
                var props = new List<Prop>();
                if (row != null)
                {
                    foreach (var c in cols)
                    {
                        if (cebOffset + c.Bit / 8 >= row.Length || (row[cebOffset + c.Bit / 8] & (1 << (7 - c.Bit % 8))) == 0) continue;
                        ushort id = (ushort)(c.Tag >> 16), type = (ushort)(c.Tag & 0xFFFF);
                        if (id == 0x67F2 || id == 0x67F3) continue; // internal row id / version columns
                        try
                        {
                            Prop p;
                            if (c.Cb == 8) { var v = new byte[8]; Buffer.BlockCopy(row, c.Ib, v, 0, 8); p = PropertyContext.Decode(heap, id, type, 0, true, v); }
                            else
                            {
                                uint raw = c.Cb == 4 ? BitConverter.ToUInt32(row, c.Ib) : c.Cb == 2 ? BitConverter.ToUInt16(row, c.Ib) : row[c.Ib];
                                p = PropertyContext.Decode(heap, id, type, raw, true);
                            }
                            if (p != null) props.Add(p);
                        }
                        catch (InvalidDataException) { }
                    }
                }
                all.Add(props);
            }
            return all;
        }
    }
}
