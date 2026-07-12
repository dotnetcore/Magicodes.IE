using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Diagnosers;
using BenchmarkDotNet.Order;
using Magicodes.IE.IO;
using MiniExcelLibs;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace Magicodes.IE.Benchmarks
{
    [MemoryDiagnoser, ThreadingDiagnoser]
    [Orderer(SummaryOrderPolicy.FastestToSlowest)]
    [CategoriesColumn]
    public class XlsxIO_Benchmarks
    {
        // ===== 1. 隔离每个 phase 测 setup vs write vs zip =====
        // 这些 micro-bench 用来定位时间损耗(setup / write rows / zip / dispose)
        // 详见 XlsxIO_Profile_Benchmarks.cs
        // ---- DTOs ----
        // [XlsxExportable] 触发源生成 fast path。Magicodes.IE.IO 用 InternalsVisibleTo
        // 授权 Magicodes.Benchmarks 访问 XlsxGeneratedGettersRegistry,benchmark csproj
        // 引用 SourceGenerator 项目作 Analyzer。生成的 *._XlsxGetters 内部类在
        // [ModuleInitializer] 注册 typed Func<object?, CellValue> getter,跳过
        // ExportProfile 反射 + BuildGettersPair 的 Expression.Compile。
        [XlsxExportable]
        public record S4(string C1, string C2, string C3, string C4);
        [XlsxExportable]
        public record N4(int Id, double Amt, long Cnt, decimal Prc);
        [XlsxExportable]
        public record D3(int Id, DateTime C1, DateTime C2, DateTime C3);
        [XlsxExportable]
        public record B3(int Id, bool B1, bool B2);
        [XlsxExportable]
        public record M5(int Id, string N, double A, DateTime D, bool B);
        [XlsxExportable]
        public record S2(string L, double V);
        // 高重复 string(SST dedupe 关键场景):4 列,总 unique 仅 16
        [XlsxExportable]
        public record R4(string A, string B, string C, string D);
        // 5 列 styled + string
        [XlsxExportable]
        public record SX5(string K, string V, string W, double P, int N);
        // Wide: 1 string key + 10 decimal value 列
        [XlsxExportable]
        public record W1K11(string K, decimal V1, decimal V2, decimal V3, decimal V4, decimal V5, decimal V6, decimal V7, decimal V8, decimal V9, decimal V10);

        // ---- data builders ----
        static S4[] S(int n) => Enumerable.Range(0, n).Select(i => new S4("s" + i, "c" + (i % 100), "p" + (i % 50), "t" + (i % 5))).ToArray();
        static N4[] Num(int n) => Enumerable.Range(0, n).Select(i => new N4(i, i * 1.5, i * 1000L, (decimal)(i * 0.99))).ToArray();
        static D3[] Dt(int n) { var b = new DateTime(2020, 1, 1); return Enumerable.Range(0, n).Select(i => new D3(i, b.AddDays(i), b.AddHours(i), b.AddMinutes(i))).ToArray(); }
        static B3[] Bo(int n) => Enumerable.Range(0, n).Select(i => new B3(i, i % 2 == 0, i % 3 == 0)).ToArray();
        static M5[] Mx(int n) { var b = new DateTime(2020, 1, 1); return Enumerable.Range(0, n).Select(i => new M5(i, "x" + i, i * 1.5, b.AddDays(i), i % 2 == 0)).ToArray(); }
        static S2[] St(int n) => Enumerable.Range(0, n).Select(i => new S2("S" + i, i * 0.1)).ToArray();
        // R4: 4 列 × 2 unique value = 16 总 unique(高 dedupe 率)
        static string[] _R4pool = { "alpha", "beta", "gamma", "delta" };
        static R4[] Rep(int n) => Enumerable.Range(0, n).Select(i => new R4(_R4pool[i % 2], _R4pool[(i / 2) % 2], _R4pool[(i / 4) % 2], _R4pool[(i / 8) % 2])).ToArray();
        static SX5[] SX(int n) => Enumerable.Range(0, n).Select(i => new SX5($"k{i % 50}", $"v{i % 100}", $"w{i % 200}", i * 0.5, i)).ToArray();
        static W1K11[] Wd(int n) => Enumerable.Range(0, n).Select(i => new W1K11($"k{i % 100}",
            (decimal)(i * 0.01), (decimal)(i * 0.02), (decimal)(i * 0.03), (decimal)(i * 0.04), (decimal)(i * 0.05),
            (decimal)(i * 0.06), (decimal)(i * 0.07), (decimal)(i * 0.08), (decimal)(i * 0.09), (decimal)(i * 0.10))).ToArray();

        // ---- fields ----
        private S4[] _1k, _10k, _50k, _100k;
        private N4[] _n, _n100k; private D3[] _d, _d100k; private B3[] _b, _b100k; private M5[] _m; private S2[] _st;
        private R4[] _r10k, _r100k; private SX5[] _sx10k; private W1K11[] _w10k;
        private S4[] _ms2A = null!;
        private S4[] _ms2B = null!;
        private Magicodes.IE.IO.Sheet[] _ms2 = null!;
        private Magicodes.IE.IO.Sheet[] _ms5 = null!;
        private Magicodes.IE.IO.Sheet[] _ms10 = null!;
        // styled profile 配置以 Action 形式复用(新 API 接受 Action<ExportProfile<T>>)
        private Action<ExportProfile<SX5>> _styledSxConfigure = null!;
        private Action<ExportProfile<S2>> _styledSConfigure = null!;
        private XlsxWriteOptions _denseSOptions = null!;
        private XlsxWriteOptions _denseSNoCompressOptions = null!;
        private XlsxWriteOptions _denseRRelaxedOptions = null!;
        private XlsxWriteOptions _denseRRelaxedNoCompressOptions = null!;

        [GlobalSetup]
        public void Setup()
        {
            _1k = S(1_000); _10k = S(10_000); _50k = S(50_000); _100k = S(100_000);
            _n = Num(10_000); _n100k = Num(100_000);
            _d = Dt(10_000); _d100k = Dt(100_000);
            _b = Bo(10_000); _b100k = Bo(100_000);
            _m = Mx(10_000); _st = St(10_000);
            _r10k = Rep(10_000); _r100k = Rep(100_000);
            _sx10k = SX(10_000); _w10k = Wd(10_000);
            _ms2A = _10k.Take(5000).ToArray();
            _ms2B = _10k.Skip(5000).ToArray();
            _ms2 = BuildMultiSheetArgs(2, 5000);
            _ms5 = BuildMultiSheetArgs(5, 2000);
            _ms10 = BuildMultiSheetArgs(10, 1000);
            // 新 API 接受 Action<ExportProfile<T>>? 而非 ExportProfile<T>
            // _reuseProfile 原为 new ExportProfile<S4>()(无配置)→ 直接用 Xlsx.ToBytes(d)
            // _sstProfile 原为 new ExportProfile<R4>().WithAutoSst(true) → 内联为 Action
            _styledSxConfigure = p =>
            {
                p.Column(x => x.K, c => c.WithBold().WithBackgroundColor("FFFCE4D6").WithBorder())
                 .Column(x => x.V, c => c.WithBackgroundColor("FFE2EFDA"))
                 .Column(x => x.W, c => c.WithItalic())
                 .Column(x => x.P, c => c.WithFormat("0.00").WithFontColor("FF0000FF"));
            };
            _styledSConfigure = p =>
            {
                p.Column(x => x.L, c => c.WithBold().WithBackgroundColor("FFFCE4D6").WithBorder())
                 .Column(x => x.V, c => c.WithFormat("0.00").WithFontColor("FF0000FF"));
            };
            _denseSOptions = new XlsxWriteOptions { StrictCellReferences = false };
            _denseSNoCompressOptions = new XlsxWriteOptions { Compression = System.IO.Compression.CompressionLevel.NoCompression, StrictCellReferences = false };
            _denseRRelaxedOptions = new XlsxWriteOptions { StrictCellReferences = false };
            _denseRRelaxedNoCompressOptions = new XlsxWriteOptions { Compression = System.IO.Compression.CompressionLevel.NoCompression, StrictCellReferences = false };
        }

        // ===== 1k-string =====
        [BenchmarkCategory("1k-s"), Benchmark] public byte[] Mio_1k_S() => Mio(_1k);
        [BenchmarkCategory("1k-s"), Benchmark] public byte[] Mio_1k_S_ProfileReuse() => Xlsx.ToBytes(_1k);

        [BenchmarkCategory("1k-s"), Benchmark] public byte[] Mini_1k_S() => Mini(_1k);
        [BenchmarkCategory("1k-s"), Benchmark] public byte[] Closed_1k_S() => Closed(_1k);
        [BenchmarkCategory("1k-s"), Benchmark] public byte[] OXml_1k_S() => OXml(_1k);
        [BenchmarkCategory("1k-s"), Benchmark] public byte[] EP_1k_S() => EP(_1k);

        // ===== 10k-string =====
        [BenchmarkCategory("10k-s"), Benchmark] public byte[] Mio_10k_S() => Mio(_10k);

        [BenchmarkCategory("10k-s"), Benchmark] public byte[] Mini_10k_S() => Mini(_10k);
        [BenchmarkCategory("10k-s"), Benchmark] public byte[] Closed_10k_S() => Closed(_10k);
        [BenchmarkCategory("10k-s"), Benchmark] public byte[] OXml_10k_S() => OXml(_10k);
        [BenchmarkCategory("10k-s"), Benchmark] public byte[] EP_10k_S() => EP(_10k);

        // ===== 100k-string =====
        [BenchmarkCategory("100k-s"), Benchmark(Baseline = true)] public byte[] Mio_100k_S() => Mio(_100k);
        [BenchmarkCategory("100k-s"), Benchmark] public byte[] Mio_100k_S_Relaxed() => Xlsx.ToBytes(_100k, options: _denseSOptions);
        [BenchmarkCategory("100k-s"), Benchmark] public byte[] Mio_100k_S_NoCompress() => Mio(_100k, System.IO.Compression.CompressionLevel.NoCompression);
        [BenchmarkCategory("100k-s"), Benchmark] public byte[] Mio_100k_S_NoCompress_Relaxed() => Xlsx.ToBytes(_100k, options: _denseSNoCompressOptions);

        [BenchmarkCategory("100k-s"), Benchmark] public byte[] Mini_100k_S() => Mini(_100k);
        [BenchmarkCategory("100k-s"), Benchmark] public byte[] Closed_100k_S() => Closed(_100k);
        [BenchmarkCategory("100k-s"), Benchmark] public byte[] EP_100k_S() => EP(_100k);

        // ===== 50k-string (中量级) =====
        [BenchmarkCategory("50k-s"), Benchmark] public byte[] Mio_50k_S() => Mio(_50k);
        [BenchmarkCategory("50k-s"), Benchmark] public byte[] Mio_50k_S_NoCompress() => Mio(_50k, System.IO.Compression.CompressionLevel.NoCompression);


        // ===== 100k-number =====
        [BenchmarkCategory("100k-n"), Benchmark] public byte[] Mio_100k_N() => Mio(_n100k);
        [BenchmarkCategory("100k-n"), Benchmark] public byte[] Mio_100k_N_Relaxed() => Xlsx.ToBytes(_n100k, options: _denseSOptions);
        [BenchmarkCategory("100k-n"), Benchmark] public byte[] Mio_100k_N_NoCompress() => Mio(_n100k, System.IO.Compression.CompressionLevel.NoCompression);
        [BenchmarkCategory("100k-n"), Benchmark] public byte[] Mio_100k_N_NoCompress_Relaxed() => Xlsx.ToBytes(_n100k, options: _denseSNoCompressOptions);

        [BenchmarkCategory("100k-n"), Benchmark] public byte[] Mini_100k_N() => Mini(_n100k);

        // ===== 100k-datetime =====
        [BenchmarkCategory("100k-d"), Benchmark] public byte[] Mio_100k_D() => Mio(_d100k);
        [BenchmarkCategory("100k-d"), Benchmark] public byte[] Mio_100k_D_Relaxed() => Xlsx.ToBytes(_d100k, options: _denseSOptions);
        [BenchmarkCategory("100k-d"), Benchmark] public byte[] Mio_100k_D_NoCompress() => Mio(_d100k, System.IO.Compression.CompressionLevel.NoCompression);
        [BenchmarkCategory("100k-d"), Benchmark] public byte[] Mio_100k_D_NoCompress_Relaxed() => Xlsx.ToBytes(_d100k, options: _denseSNoCompressOptions);

        [BenchmarkCategory("100k-d"), Benchmark] public byte[] Mini_100k_D() => Mini(_d100k);

        // ===== 100k-boolean =====
        [BenchmarkCategory("100k-b"), Benchmark] public byte[] Mio_100k_B() => Mio(_b100k);
        [BenchmarkCategory("100k-b"), Benchmark] public byte[] Mio_100k_B_Relaxed() => Xlsx.ToBytes(_b100k, options: _denseSOptions);
        [BenchmarkCategory("100k-b"), Benchmark] public byte[] Mio_100k_B_NoCompress() => Mio(_b100k, System.IO.Compression.CompressionLevel.NoCompression);
        [BenchmarkCategory("100k-b"), Benchmark] public byte[] Mio_100k_B_NoCompress_Relaxed() => Xlsx.ToBytes(_b100k, options: _denseSNoCompressOptions);

        [BenchmarkCategory("100k-b"), Benchmark] public byte[] Mini_100k_B() => Mini(_b100k);

        // ===== SST 场景:高重复 string =====
        [BenchmarkCategory("sst-10k"), Benchmark] public byte[] Mio_10k_Repeated() => Mio(_r10k);
        [BenchmarkCategory("sst-10k"), Benchmark] public byte[] Mio_10k_Repeated_AutoSst() => Xlsx.ToBytes(_r10k, p => p.WithAutoSst(true));

        [BenchmarkCategory("sst-10k"), Benchmark] public byte[] Mini_10k_Repeated() => Mini(_r10k);
        [BenchmarkCategory("sst-100k"), Benchmark] public byte[] Mio_100k_Repeated() => Mio(_r100k);
        [BenchmarkCategory("sst-100k"), Benchmark] public byte[] Mio_100k_Repeated_Relaxed() => Xlsx.ToBytes(_r100k, options: _denseRRelaxedOptions);
        [BenchmarkCategory("sst-100k"), Benchmark] public byte[] Mio_100k_Repeated_NoCompress() => Mio(_r100k, System.IO.Compression.CompressionLevel.NoCompression);
        [BenchmarkCategory("sst-100k"), Benchmark] public byte[] Mio_100k_Repeated_NoCompress_Relaxed() => Xlsx.ToBytes(_r100k, options: _denseRRelaxedNoCompressOptions);
        [BenchmarkCategory("sst-100k"), Benchmark] public byte[] Mio_100k_Repeated_AutoSst() => Xlsx.ToBytes(_r100k, p => p.WithAutoSst(true));
        [BenchmarkCategory("sst-100k"), Benchmark] public byte[] Mio_100k_Repeated_AutoSst_Relaxed() => Xlsx.ToBytes(_r100k, p => p.WithAutoSst(true), _denseRRelaxedOptions);
        [BenchmarkCategory("sst-100k"), Benchmark] public byte[] Mio_100k_Repeated_AutoSst_NoCompress() => Xlsx.ToBytes(_r100k, p => p.WithAutoSst(true), new XlsxWriteOptions { Compression = System.IO.Compression.CompressionLevel.NoCompression });
        [BenchmarkCategory("sst-100k"), Benchmark] public byte[] Mio_100k_Repeated_AutoSst_NoCompress_Relaxed() => Xlsx.ToBytes(_r100k, p => p.WithAutoSst(true), _denseRRelaxedNoCompressOptions);

        [BenchmarkCategory("sst-100k"), Benchmark] public byte[] Mini_100k_Repeated() => Mini(_r100k);

        // ===== Multi-sheet 多档 =====
        [BenchmarkCategory("ms2"), Benchmark] public byte[] Mio_MS2() => Xlsx.WriteWorkbookToBytes(_ms2);
        [BenchmarkCategory("ms5"), Benchmark] public byte[] Mio_MS5() => Xlsx.WriteWorkbookToBytes(_ms5);
        [BenchmarkCategory("ms10"), Benchmark] public byte[] Mio_MS10() => Xlsx.WriteWorkbookToBytes(_ms10);
        [BenchmarkCategory("ms2"), Benchmark] public byte[] EP_MS2() { using var ms = new MemoryStream(); using var p = new OfficeOpenXml.ExcelPackage(ms); EPWrite(p, "A", _ms2A); EPWrite(p, "B", _ms2B); p.Save(); return ms.ToArray(); }

        // ===== Wide (1 key + 10 decimal) =====
        [BenchmarkCategory("wide-10k"), Benchmark] public byte[] Mio_10k_Wide() => Mio(_w10k);

        [BenchmarkCategory("wide-10k"), Benchmark] public byte[] Mini_10k_Wide() => Mini(_w10k);

        // ===== Styled + string(5 列) =====
        [BenchmarkCategory("10k-sx"), Benchmark] public byte[] Mio_10k_SX() => MioStyledSX(_sx10k);
        [BenchmarkCategory("10k-sx"), Benchmark] public byte[] Mio_10k_SX_FastPath() => Mio(_sx10k);

        // ===== 10k-number =====
        [BenchmarkCategory("10k-n"), Benchmark] public byte[] Mio_10k_N() => Mio(_n);

        [BenchmarkCategory("10k-n"), Benchmark] public byte[] Mini_10k_N() => Mini(_n);
        [BenchmarkCategory("10k-n"), Benchmark] public byte[] Closed_10k_N() => Closed(_n);
        [BenchmarkCategory("10k-n"), Benchmark] public byte[] EP_10k_N() => EP(_n);

        // ===== 10k-datetime =====
        [BenchmarkCategory("10k-d"), Benchmark] public byte[] Mio_10k_D() => Mio(_d);

        [BenchmarkCategory("10k-d"), Benchmark] public byte[] Mini_10k_D() => Mini(_d);
        [BenchmarkCategory("10k-d"), Benchmark] public byte[] Closed_10k_D() => Closed(_d);
        [BenchmarkCategory("10k-d"), Benchmark] public byte[] EP_10k_D() => EP(_d);

        // ===== 10k-boolean =====
        [BenchmarkCategory("10k-b"), Benchmark] public byte[] Mio_10k_B() => Mio(_b);

        [BenchmarkCategory("10k-b"), Benchmark] public byte[] Mini_10k_B() => Mini(_b);
        [BenchmarkCategory("10k-b"), Benchmark] public byte[] Closed_10k_B() => Closed(_b);
        [BenchmarkCategory("10k-b"), Benchmark] public byte[] EP_10k_B() => EP(_b);

        // ===== 10k-mixed =====
        [BenchmarkCategory("10k-m"), Benchmark] public byte[] Mio_10k_M() => Mio(_m);

        [BenchmarkCategory("10k-m"), Benchmark] public byte[] Mini_10k_M() => Mini(_m);
        [BenchmarkCategory("10k-m"), Benchmark] public byte[] Closed_10k_M() => Closed(_m);
        [BenchmarkCategory("10k-m"), Benchmark] public byte[] EP_10k_M() => EP(_m);

        // ===== 10k-styled =====
        [BenchmarkCategory("10k-st"), Benchmark] public byte[] Mio_10k_St() => MioStyled(_st);
        [BenchmarkCategory("10k-st"), Benchmark] public byte[] Closed_10k_St() => ClosedStyled(_st);
        [BenchmarkCategory("10k-st"), Benchmark] public byte[] EP_10k_St() => EPStyled(_st);

        // ===== multi-sheet =====
        // ===== implementation =====

        static byte[] Mio<T>(T[] d, System.IO.Compression.CompressionLevel compression = System.IO.Compression.CompressionLevel.Fastest) where T : class => Xlsx.ToBytes(d, options: new XlsxWriteOptions { Compression = compression });

        // 多 sheet 工厂
        static Magicodes.IE.IO.Sheet[] BuildMultiSheetArgs(int sheetCount, int rowsPerSheet)
        {
            var args = new Magicodes.IE.IO.Sheet[sheetCount];
            for (int s = 0; s < sheetCount; s++)
            {
                var arr = new S4[rowsPerSheet];
                for (int i = 0; i < rowsPerSheet; i++)
                    arr[i] = new S4($"s{s}r{i}", $"c{i % 100}", $"p{i % 50}", $"t{i % 5}");
                args[s] = new Magicodes.IE.IO.Sheet($"S{s}", arr);
            }
            return args;
        }

        // 5 列 styled + string
        byte[] MioStyledSX(SX5[] d) => Xlsx.ToBytes(d, _styledSxConfigure);

        byte[] MioStyled(S2[] d) => Xlsx.ToBytes(d, _styledSConfigure);


        static byte[] Mini<T>(T[] d)
        {
            using var ms = new MemoryStream();
            MiniExcelLibs.MiniExcel.SaveAs(ms, d, true, "S1", ExcelType.XLSX);
            return ms.ToArray();
        }

        static byte[] Closed<T>(T[] d)
        {
            using var wb = new XLWorkbook();
            wb.Worksheets.Add("S1").Cell(1, 1).InsertData(d);
            using var ms = new MemoryStream(); wb.SaveAs(ms); return ms.ToArray();
        }

        static byte[] ClosedStyled(S2[] d)
        {
            using var wb = new XLWorkbook(); var ws = wb.Worksheets.Add("S1");
            ws.Cell(1, 1).Value = "L"; ws.Cell(1, 1).Style.Font.Bold = true; ws.Cell(1, 1).Style.Fill.BackgroundColor = XLColor.FromArgb(252, 228, 214); ws.Cell(1, 1).Style.Border.BottomBorder = XLBorderStyleValues.Thin;
            ws.Cell(1, 2).Value = "V"; ws.Cell(1, 2).Style.Font.FontColor = XLColor.Blue;
            for (int i = 0; i < d.Length; i++) { ws.Cell(i + 2, 1).Value = d[i].L; ws.Cell(i + 2, 2).Value = d[i].V; }
            ws.Range(2, 2, d.Length + 1, 2).Style.NumberFormat.Format = "0.00";
            using var ms = new MemoryStream(); wb.SaveAs(ms); return ms.ToArray();
        }

        static byte[] OXml<T>(T[] d)
        {
            using var ms = new MemoryStream();
            using var doc = SpreadsheetDocument.Create(ms, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook);
            var wbp = doc.AddWorkbookPart(); wbp.Workbook = new Workbook();
            var sp = wbp.AddNewPart<WorksheetPart>(); sp.Worksheet = new Worksheet(new SheetData());
            wbp.Workbook.AppendChild(new Sheets()).Append(new DocumentFormat.OpenXml.Spreadsheet.Sheet { Id = wbp.GetIdOfPart(sp), SheetId = 1, Name = "S1" });
            var sd = sp.Worksheet.GetFirstChild<SheetData>()!;
            var pr = typeof(T).GetProperties();
            var h = new Row(); foreach (var p in pr) h.AppendChild(new Cell { DataType = CellValues.String, CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue(p.Name) }); sd.AppendChild(h);
            foreach (var item in d)
            {
                var row = new Row();
                foreach (var p in pr)
                {
                    var v = p.GetValue(item);
                    DocumentFormat.OpenXml.Spreadsheet.CellValue cv;
                    if (v is null) cv = new DocumentFormat.OpenXml.Spreadsheet.CellValue("");
                    else if (v is string sv) cv = new DocumentFormat.OpenXml.Spreadsheet.CellValue(sv);
                    else if (v is int iv) cv = new DocumentFormat.OpenXml.Spreadsheet.CellValue((double)iv);
                    else if (v is long lv) cv = new DocumentFormat.OpenXml.Spreadsheet.CellValue((double)lv);
                    else if (v is double dv) cv = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dv);
                    else if (v is decimal mv) cv = new DocumentFormat.OpenXml.Spreadsheet.CellValue((double)mv);
                    else if (v is DateTime dtv) cv = new DocumentFormat.OpenXml.Spreadsheet.CellValue(dtv.ToOADate().ToString(CultureInfo.InvariantCulture));
                    else if (v is bool bv) cv = new DocumentFormat.OpenXml.Spreadsheet.CellValue(bv ? "1" : "0");
                    else cv = new DocumentFormat.OpenXml.Spreadsheet.CellValue(v.ToString() ?? "");
                    row.AppendChild(new Cell { DataType = CellValues.String, CellValue = cv });
                }
                sd.AppendChild(row);
            }
            wbp.Workbook.Save(); return ms.ToArray();
        }

        static byte[] EP<T>(T[] d)
        {
            var pr = typeof(T).GetProperties();
            using var ms = new MemoryStream();
            using var pkg = new OfficeOpenXml.ExcelPackage(ms);
            var ws = pkg.Workbook.Worksheets.Add("S1");
            for (int c = 0; c < pr.Length; c++) ws.Cells[1, c + 1].Value = pr[c].Name;
            for (int r = 0; r < d.Length; r++) for (int c = 0; c < pr.Length; c++) ws.Cells[r + 2, c + 1].Value = pr[c].GetValue(d[r]);
            pkg.Save(); return ms.ToArray();
        }

        static byte[] EPStyled(S2[] d)
        {
            using var ms = new MemoryStream(); using var pkg = new OfficeOpenXml.ExcelPackage(ms); var ws = pkg.Workbook.Worksheets.Add("S1");
            ws.Cells[1, 1].Value = "L"; ws.Cells[1, 1].Style.Font.Bold = true; ws.Cells[1, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid; ws.Cells[1, 1].Style.Fill.BackgroundColor.SetColor(SixLabors.ImageSharp.Color.FromRgba(252, 228, 214, 255)); ws.Cells[1, 1].Style.Border.BorderAround(OfficeOpenXml.Style.ExcelBorderStyle.Thin);
            ws.Cells[1, 2].Value = "V"; ws.Cells[1, 2].Style.Font.Color.SetColor(SixLabors.ImageSharp.Color.FromRgb(0, 0, 255));
            for (int i = 0; i < d.Length; i++) { ws.Cells[i + 2, 1].Value = d[i].L; ws.Cells[i + 2, 2].Value = d[i].V; }
            ws.Cells[2, 2, d.Length + 1, 2].Style.Numberformat.Format = "0.00";
            pkg.Save(); return ms.ToArray();
        }

        static void EPWrite<T>(OfficeOpenXml.ExcelPackage pkg, string name, T[] d)
        {
            var ws = pkg.Workbook.Worksheets.Add(name); var pr = typeof(T).GetProperties();
            for (int c = 0; c < pr.Length; c++) ws.Cells[1, c + 1].Value = pr[c].Name;
            for (int r = 0; r < d.Length; r++) for (int c = 0; c < pr.Length; c++) ws.Cells[r + 2, c + 1].Value = pr[c].GetValue(d[r]);
        }
    }
}
