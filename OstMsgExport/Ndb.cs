// Node database layer of Outlook OST/PST files ([MS-PST] 2.2), read-only.
// Supports Unicode PST/OST (version 23, 512-byte pages) and Unicode 4K OST (version 36, 4096-byte pages with
// zlib-compressed blocks, written by Outlook 2013 and later). The file is opened read-only and never changed.
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;

namespace OstMsgExport
{
    sealed class NbtEntry
    {
        public uint Nid; public ulong BidData; public ulong BidSub; public uint NidParent;
        public int Type => (int)(Nid & 0x1F);
    }

    sealed class BbtEntry
    {
        public ulong Bid; public long Ib; public int Cb; public int CbInflated;
    }

    sealed class OstFile : IDisposable
    {
        readonly FileStream fs;
        readonly object gate = new object();
        public readonly int Version;
        public readonly bool Is4K;
        readonly int pageSize;
        public readonly Dictionary<uint, NbtEntry> Nodes = new Dictionary<uint, NbtEntry>();
        public readonly Dictionary<ulong, BbtEntry> Blocks = new Dictionary<ulong, BbtEntry>();
        public readonly string Path;
        public readonly long Length;

        public OstFile(string path)
        {
            Path = path;
            fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, 1 << 16);
            Length = fs.Length;
            var h = Read(0, 600);
            if (h[0] != 0x21 || h[1] != 0x42 || h[2] != 0x44 || h[3] != 0x4E) throw new InvalidDataException("This is not an Outlook data file (OST/PST): the file signature is wrong.");
            Version = BitConverter.ToUInt16(h, 10);
            if (Version == 14 || Version == 15) throw new InvalidDataException("This is an old ANSI Outlook file (Outlook 2002 or earlier). It is not supported.");
            if (Version != 23 && Version != 36) throw new InvalidDataException($"Unsupported Outlook file version {Version}.");
            Is4K = Version == 36;
            pageSize = Is4K ? 4096 : 512;
            if (h[513] != 0) throw new InvalidDataException($"This file uses encryption method {h[513]}, which is not supported yet.");
            long nbtIb = BitConverter.ToInt64(h, 224), bbtIb = BitConverter.ToInt64(h, 240);
            LoadTree(bbtIb, 0x80, 0);
            LoadTree(nbtIb, 0x81, 0);
        }

        public byte[] Read(long ib, int n)
        {
            lock (gate)
            {
                var b = new byte[n];
                fs.Position = ib;
                int got = 0;
                while (got < n) { int r = fs.Read(b, got, n - got); if (r <= 0) throw new EndOfStreamException($"Unexpected end of file at {ib + got}."); got += r; }
                return b;
            }
        }

        void LoadTree(long ib, byte ptype, int depth)
        {
            if (depth > 16) throw new InvalidDataException("The index tree is too deep; the file is damaged.");
            if (ib <= 0 || ib + pageSize > Length) throw new InvalidDataException($"An index page points outside the file ({ib}).");
            var p = Read(ib, pageSize);
            int cEnt, cbEnt, level, ptypeAt;
            if (Is4K) { cEnt = BitConverter.ToUInt16(p, 4056); cbEnt = p[4060]; level = p[4061]; ptypeAt = 4072; }
            else { cEnt = p[488]; cbEnt = p[490]; level = p[491]; ptypeAt = 496; }
            if (p[ptypeAt] != ptype) throw new InvalidDataException($"Index page at {ib} has type 0x{p[ptypeAt]:X2}, expected 0x{ptype:X2}; the file is damaged.");
            int max = (Is4K ? 4056 : 488) / Math.Max(1, cbEnt);
            if (cEnt > max) throw new InvalidDataException($"Index page at {ib} is damaged.");
            for (int i = 0; i < cEnt; i++)
            {
                int o = i * cbEnt;
                if (level > 0) { LoadTree(BitConverter.ToInt64(p, o + 16), ptype, depth + 1); continue; }
                if (ptype == 0x80)
                {
                    var e = new BbtEntry { Bid = BitConverter.ToUInt64(p, o) & ~1UL, Ib = BitConverter.ToInt64(p, o + 8), Cb = BitConverter.ToUInt16(p, o + 16) };
                    e.CbInflated = Is4K ? BitConverter.ToUInt16(p, o + 18) : e.Cb;
                    Blocks[e.Bid] = e;
                }
                else
                {
                    var e = new NbtEntry { Nid = (uint)BitConverter.ToUInt64(p, o), BidData = BitConverter.ToUInt64(p, o + 8), BidSub = BitConverter.ToUInt64(p, o + 16), NidParent = BitConverter.ToUInt32(p, o + 24) };
                    Nodes[e.Nid] = e;
                }
            }
        }

        /// <summary>Raw contents of one block, decompressed.</summary>
        public byte[] ReadBlock(ulong bid)
        {
            if (!Blocks.TryGetValue(bid & ~1UL, out var e)) throw new InvalidDataException($"Block {bid} is missing from the block index.");
            if (e.Ib <= 0 || e.Ib + e.Cb > Length) throw new InvalidDataException($"Block {bid} points outside the file.");
            var raw = Read(e.Ib, e.Cb);
            if (!Is4K || e.CbInflated == e.Cb || e.CbInflated == 0) return raw;
            // Compressed with zlib: a 2-byte header, raw deflate data, then an Adler-32 checksum.
            if (raw.Length < 2 || (raw[0] & 0x0F) != 8) throw new InvalidDataException($"Block {bid} is not valid compressed data.");
            var outBuf = new byte[e.CbInflated];
            using (var ds = new DeflateStream(new MemoryStream(raw, 2, raw.Length - 2), CompressionMode.Decompress))
            {
                int got = 0;
                while (got < outBuf.Length) { int r = ds.Read(outBuf, got, outBuf.Length - got); if (r <= 0) break; got += r; }
                if (got != outBuf.Length) throw new InvalidDataException($"Block {bid} decompressed to {got} bytes, expected {outBuf.Length}.");
            }
            return outBuf;
        }

        static bool IsInternal(ulong bid) => (bid & 2) != 0;

        /// <summary>The data blocks of a node, in order (an XBLOCK/XXBLOCK tree is flattened).</summary>
        public List<byte[]> ReadDataBlocks(ulong bid)
        {
            var list = new List<byte[]>();
            if (bid == 0) return list;
            AddData(bid, list, 0);
            return list;
        }

        void AddData(ulong bid, List<byte[]> list, int depth)
        {
            if (depth > 4) throw new InvalidDataException("Data tree too deep.");
            var b = ReadBlock(bid);
            if (!IsInternal(bid)) { list.Add(b); return; }
            if (b[0] != 0x01) throw new InvalidDataException($"Block {bid} should be a data tree block but is type {b[0]}.");
            int cEnt = BitConverter.ToUInt16(b, 2);
            for (int i = 0; i < cEnt; i++) AddData(BitConverter.ToUInt64(b, 8 + i * 8), list, depth + 1);
        }

        /// <summary>The subnodes of a node: nid -> (data bid, subnode bid).</summary>
        public Dictionary<uint, NbtEntry> ReadSubnodes(ulong bidSub)
        {
            var d = new Dictionary<uint, NbtEntry>();
            if (bidSub != 0) AddSub(bidSub, d, 0);
            return d;
        }

        void AddSub(ulong bid, Dictionary<uint, NbtEntry> d, int depth)
        {
            if (depth > 4) throw new InvalidDataException("Subnode tree too deep.");
            var b = ReadBlock(bid);
            if (b[0] != 0x02) throw new InvalidDataException($"Block {bid} should be a subnode block but is type {b[0]}.");
            int level = b[1], cEnt = BitConverter.ToUInt16(b, 2);
            for (int i = 0; i < cEnt; i++)
            {
                if (level == 0)
                {
                    int o = 8 + i * 24;
                    var e = new NbtEntry { Nid = (uint)BitConverter.ToUInt64(b, o), BidData = BitConverter.ToUInt64(b, o + 8), BidSub = BitConverter.ToUInt64(b, o + 16) };
                    d[e.Nid] = e;
                }
                else AddSub(BitConverter.ToUInt64(b, 8 + i * 16 + 8), d, depth + 1);
            }
        }

        public void Dispose() => fs.Dispose();
    }

    /// <summary>A node (top-level or subnode) with lazily loaded data blocks and subnodes.</summary>
    sealed class Node
    {
        readonly OstFile ost;
        public readonly uint Nid;
        readonly ulong bidData, bidSub;
        List<byte[]> blocks;
        Dictionary<uint, NbtEntry> subs;

        public Node(OstFile ost, uint nid, ulong bidData, ulong bidSub) { this.ost = ost; Nid = nid; this.bidData = bidData; this.bidSub = bidSub; }
        public Node(OstFile ost, NbtEntry e) : this(ost, e.Nid, e.BidData, e.BidSub) { }

        public List<byte[]> Blocks => blocks ?? (blocks = ost.ReadDataBlocks(bidData));
        public bool Is4K => ost.Is4K;

        public byte[] Data()
        {
            var bl = Blocks;
            if (bl.Count == 1) return bl[0];
            int n = 0; foreach (var b in bl) n += b.Length;
            var r = new byte[n]; int o = 0;
            foreach (var b in bl) { Buffer.BlockCopy(b, 0, r, o, b.Length); o += b.Length; }
            return r;
        }

        public Node Subnode(uint nid)
        {
            if (subs == null) subs = ost.ReadSubnodes(bidSub);
            return subs.TryGetValue(nid, out var e) ? new Node(ost, e) : null;
        }
    }
}
