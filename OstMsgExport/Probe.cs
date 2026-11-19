using System;
using System.IO;
using System.Linq;

namespace OstMsgExport
{
    // Diagnostic dump of the low-level layout (development aid: OstMsgExport.exe --probe <file>).
    static class Probe
    {
        static string Hex(byte[] b, int s, int n) => string.Join(" ", b.Skip(s).Take(n).Select(x => x.ToString("X2")));

        public static void Run(string path)
        {
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                byte[] Read(long ib, int n) { var b = new byte[n]; fs.Position = ib; fs.Read(b, 0, n); return b; }
                var h = Read(0, 600);
                long nbt = BitConverter.ToInt64(h, 224), bbt = BitConverter.ToInt64(h, 240);
                byte[] leaf(long ib, string name)
                {
                    while (true)
                    {
                        var p = Read(ib, 4096);
                        int c = BitConverter.ToUInt16(p, 4056), cb = p[4060], lvl = p[4061];
                        Console.WriteLine($"{name} page @{ib} cEnt={c} cbEnt={cb} level={lvl} ptype={p[4072]:X2}");
                        if (lvl == 0) { for (int i = 0; i < Math.Min(c, 4); i++) Console.WriteLine("  " + Hex(p, i * cb, cb)); return p; }
                        ib = BitConverter.ToInt64(p, 16);
                    }
                }
                var bl = leaf(bbt, "BBT");
                leaf(nbt, "NBT");
                for (int i = 0; i < 4; i++)
                {
                    long ib = BitConverter.ToInt64(bl, i * 24 + 8);
                    int cb = BitConverter.ToUInt16(bl, i * 24 + 16), cb2 = BitConverter.ToUInt16(bl, i * 24 + 18);
                    var blk = Read(ib, 512 * 20);
                    Console.WriteLine($"block {i} ib={ib} cb={cb} f18={cb2} start: {Hex(blk, 0, 16)}");
                    // find trailer: bid echo after data
                    ulong bid = BitConverter.ToUInt64(bl, i * 24);
                    for (int o = cb; o < blk.Length - 8; o++)
                        if (BitConverter.ToUInt64(blk, o) == bid) { Console.WriteLine($"   bid echo at +{o}: {Hex(blk, o - 8, 32)}"); break; }
                }
            }
        }
    }
}
