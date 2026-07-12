// ======================================================================
//  XlsxIO_Profile_Benchmarks.cs — 隔离每个 phase 测损耗
//
//  目的:把 "100k string 24ms" 拆成 setup / write / zip / dispose 4 段,
//  定位时间花在哪。
//
//  设计:
//  - Mio_Phase1_Setup: new writer + AddSheet + WriteSheetMeta + WriteHeader(无 row)
//  - Mio_Phase2_Write: 复用已 AddSheet 的 writer,只 WriteRows 100k
//  - Mio_Phase3_Dispose: 复用已写完的 writer,只 Dispose
//  - Mio_Full: 全流程(= phase1+2+3)
//
//  用 [InvocationCount(1)] 避免 BDN 默认 16 次 invocation 抹掉 setup 开销。
// ======================================================================

#nullable enable

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Order;
using Magicodes.IE.IO;

namespace Magicodes.IE.Benchmarks
{
    [MemoryDiagnoser]
    [Orderer(SummaryOrderPolicy.FastestToSlowest)]
    [CategoriesColumn]
    [InvocationCount(1)]
    public class XlsxIO_Profile_Benchmarks
    {
        public record S4(string C1, string C2, string C3, string C4);

        static S4[] S(int n) => System.Linq.Enumerable.Range(0, n)
            .Select(i => new S4("s" + i, "c" + (i % 100), "p" + (i % 50), "t" + (i % 5))).ToArray();

        private static readonly ColumnMeta[] s_cols = XlsxIO_TestSupport_Profile.MakeCols();
        private static readonly TypedRowPlan<S4> s_plan = XlsxIO_TestSupport_Profile.MakeTypedPlan<S4>(new[] { "C1", "C2", "C3", "C4" });

        private S4[] _10k = null!, _100k = null!;
        private System.IO.Compression.CompressionLevel _comp = System.IO.Compression.CompressionLevel.Fastest;

        [GlobalSetup]
        public void Setup()
        {
            _10k = S(10_000);
            _100k = S(100_000);
        }

        // ===== 10k 拆 phase =====

        [BenchmarkCategory("phase-10k"), Benchmark(Baseline = true)]
        public byte[] Mio_10k_Full() => MioBytes(_10k, _comp);

        [BenchmarkCategory("phase-10k"), Benchmark]
        public long Mio_10k_Full_WriteOnly() => MioLength(_10k, _comp);

        [BenchmarkCategory("phase-10k"), Benchmark]
        public int Mio_10k_Full_CopySizeOnly() => MioCopySize(_10k, _comp);

        [BenchmarkCategory("phase-10k"), Benchmark]
        public byte[] Mio_10k_Phase1_Setup()
        {
            // setup: new writer + AddSheet + WriteSheetMeta + WriteHeader
            // (不写 row, 立即 dispose 触发 zip close)
            using var ms = new MemoryStream();
            using (var w = new XlsxWriter(ms, compression: _comp))
            {
                w.AddSheet("S1");
                PrepareWriter(w);
                w.WriteSheetMeta(s_cols, freezeHeader: true);
                w.WriteHeader(s_cols);
            }
            return ms.ToArray();
        }

        [BenchmarkCategory("phase-10k"), Benchmark]
        public byte[] Mio_10k_Phase2_WriteOnly()
        {
            // 只写 rows: 复用现成 reader + cols
            // writer + setup 仍要 1 次, 但我们只测增量
            using var ms = new MemoryStream();
            using (var w = new XlsxWriter(ms, compression: _comp))
            {
                w.AddSheet("S1");
                PrepareWriter(w);
                w.WriteSheetMeta(s_cols, freezeHeader: true);
                w.WriteHeader(s_cols);
                w.WriteRows(_10k, s_plan);
            }
            return ms.ToArray();
        }

        // ===== 100k 拆 phase =====

        [BenchmarkCategory("phase-100k"), Benchmark]
        public byte[] Mio_100k_Full() => MioBytes(_100k, _comp);

        [BenchmarkCategory("phase-100k"), Benchmark]
        public long Mio_100k_Full_WriteOnly() => MioLength(_100k, _comp);

        [BenchmarkCategory("phase-100k"), Benchmark]
        public int Mio_100k_Full_CopySizeOnly() => MioCopySize(_100k, _comp);

        [BenchmarkCategory("phase-100k"), Benchmark]
        public long Mio_100k_Full_NullStream() => MioNull(_100k, _comp);

        [BenchmarkCategory("phase-100k"), Benchmark]
        public byte[] Mio_100k_Phase1_Setup()
        {
            using var ms = new MemoryStream();
            using (var w = new XlsxWriter(ms, compression: _comp))
            {
                w.AddSheet("S1");
                PrepareWriter(w);
                w.WriteSheetMeta(s_cols, freezeHeader: true);
                w.WriteHeader(s_cols);
            }
            return ms.ToArray();
        }

        [BenchmarkCategory("phase-100k"), Benchmark]
        public byte[] Mio_100k_NoCompress_Full()
        {
            using var ms = new MemoryStream();
            using (var w = new XlsxWriter(ms, compression: System.IO.Compression.CompressionLevel.NoCompression))
            {
                w.AddSheet("S1");
                PrepareWriter(w);
                w.WriteSheetMeta(s_cols, freezeHeader: true);
                w.WriteHeader(s_cols);
                w.WriteRows(_100k, s_plan);
            }
            return ms.ToArray();
        }

        [BenchmarkCategory("phase-100k"), Benchmark]
        public byte[] Mio_100k_NoCompress_Phase1_Setup()
        {
            using var ms = new MemoryStream();
            using (var w = new XlsxWriter(ms, compression: System.IO.Compression.CompressionLevel.NoCompression))
            {
                w.AddSheet("S1");
                PrepareWriter(w);
                w.WriteSheetMeta(s_cols, freezeHeader: true);
                w.WriteHeader(s_cols);
            }
            return ms.ToArray();
        }

        // ===== 2. SST path(高重复 → 走 <c t="s">) =====
        // 看 SST 路径在 100k string 上比 inline 慢还是快(escape 工作应大幅减少)
        [BenchmarkCategory("phase-100k"), Benchmark]
        public byte[] Mio_100k_SST_Full()
        {
            // _100k 4 列 unique = 4 unique count, 高度重复
            using var ms = new MemoryStream();
            using (var w = new XlsxWriter(ms))
            {
                w.AddSheet("S1");
                PrepareWriter(w);
                w.EnableSharedStrings();
                w.WriteSheetMeta(s_cols, freezeHeader: true);
                w.WriteHeader(s_cols);
                w.WriteRows(_100k, s_plan);
            }
            return ms.ToArray();
        }

        [BenchmarkCategory("phase-100k"), Benchmark]
        public byte[] Mio_100k_SST_NoCompress_Full()
        {
            using var ms = new MemoryStream();
            using (var w = new XlsxWriter(ms, compression: System.IO.Compression.CompressionLevel.NoCompression))
            {
                w.AddSheet("S1");
                PrepareWriter(w);
                w.EnableSharedStrings();
                w.WriteSheetMeta(s_cols, freezeHeader: true);
                w.WriteHeader(s_cols);
                w.WriteRows(_100k, s_plan);
            }
            return ms.ToArray();
        }

        // ===== 3. 10k 写 detail (多 phase) =====
        [BenchmarkCategory("phase-10k"), Benchmark]
        public byte[] Mio_10k_NoCompress_Full()
        {
            using var ms = new MemoryStream();
            using (var w = new XlsxWriter(ms, compression: System.IO.Compression.CompressionLevel.NoCompression))
            {
                w.AddSheet("S1");
                PrepareWriter(w);
                w.WriteSheetMeta(s_cols, freezeHeader: true);
                w.WriteHeader(s_cols);
                w.WriteRows(_10k, s_plan);
            }
            return ms.ToArray();
        }

        private static byte[] MioBytes(S4[] d, System.IO.Compression.CompressionLevel compression)
        {
            using var ms = new MemoryStream();
            using (var w = new XlsxWriter(ms, compression: compression))
            {
                w.AddSheet("S1");
                PrepareWriter(w);
                w.WriteSheetMeta(s_cols, freezeHeader: true);
                w.WriteHeader(s_cols);
                w.WriteRows(d, s_plan);
            }
            return ms.ToArray();
        }

        private static long MioLength(S4[] d, System.IO.Compression.CompressionLevel compression)
        {
            using var ms = new MemoryStream();
            using (var w = new XlsxWriter(ms, compression: compression))
            {
                w.AddSheet("S1");
                PrepareWriter(w);
                w.WriteSheetMeta(s_cols, freezeHeader: true);
                w.WriteHeader(s_cols);
                w.WriteRows(d, s_plan);
            }
            return ms.Length;
        }

        private static int MioCopySize(S4[] d, System.IO.Compression.CompressionLevel compression)
        {
            using var ms = new MemoryStream();
            using (var w = new XlsxWriter(ms, compression: compression))
            {
                w.AddSheet("S1");
                PrepareWriter(w);
                w.WriteSheetMeta(s_cols, freezeHeader: true);
                w.WriteHeader(s_cols);
                w.WriteRows(d, s_plan);
            }
            return ms.ToArray().Length;
        }

        private static long MioNull(S4[] d, System.IO.Compression.CompressionLevel compression)
        {
            using var w = new XlsxWriter(Stream.Null, compression: compression);
            w.AddSheet("S1");
            PrepareWriter(w);
            w.WriteSheetMeta(s_cols, freezeHeader: true);
            w.WriteHeader(s_cols);
            w.WriteRows(d, s_plan);
            return 0;
        }

        // ===== helpers =====
        private static void PrepareWriter(XlsxWriter writer)
        {
            writer.SetNumFmts(s_plan.BuildNumFmts());
            writer.ResolveColumnStyles(s_plan.Columns);
        }
    }

    internal static class XlsxIO_TestSupport_Profile
    {
        public static ColumnMeta[] MakeCols() => new[]
        {
            new ColumnMeta("C1", "C1", null, null, false, 0, 0),
            new ColumnMeta("C2", "C2", null, null, false, 0, 1),
            new ColumnMeta("C3", "C3", null, null, false, 0, 2),
            new ColumnMeta("C4", "C4", null, null, false, 0, 3),
        };

        public static TypedRowPlan<T> MakeTypedPlan<T>(string[] propertyNames)
        {
            var cols = new ColumnMeta[propertyNames.Length];
            for (int i = 0; i < propertyNames.Length; i++)
                cols[i] = new ColumnMeta(propertyNames[i], propertyNames[i], null, null, false, 0, i);
            var props = typeof(T).GetProperties();
            var getters = new Func<T, CellValue>[propertyNames.Length];
            for (int i = 0; i < propertyNames.Length; i++)
            {
                var p = System.Array.Find(props, x => x.Name == propertyNames[i])!;
                getters[i] = o => p.GetValue(o) switch
                {
                    string s => CellValue.FromString(s),
                    int v => CellValue.FromInteger(v),
                    _ => CellValue.Null
                };
            }
            return new TypedRowPlan<T>(cols, new Func<object?, CellValue>[0], getters, new int[propertyNames.Length], new Action<XlsxWriter.XlsxRowWriter, T, int>?[propertyNames.Length], false);
        }
    }
}
