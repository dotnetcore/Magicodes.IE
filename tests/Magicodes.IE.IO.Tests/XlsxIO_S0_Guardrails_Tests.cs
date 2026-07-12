
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using Magicodes.IE.IO;
using Shouldly;
using Xunit;

namespace Magicodes.IE.IO.Tests
{
    public partial class XlsxIO_Tests
    {

        public static IEnumerable<object[]> ColumnLetterCases()
        {
            yield return new object[] { 0, "A" };
            yield return new object[] { 25, "Z" };
            yield return new object[] { 26, "AA" };
            yield return new object[] { 27, "AB" };
            yield return new object[] { 51, "AZ" };
            yield return new object[] { 52, "BA" };
            yield return new object[] { 701, "ZZ" };
            yield return new object[] { 702, "AAA" };
            yield return new object[] { 703, "AAB" };
            yield return new object[] { 16383, "XFD" };
        }

        [Theory]
        [MemberData(nameof(ColumnLetterCases))]
        public void ColumnLetter_KnownColumns_MatchExpected(int col0, string expected)
        {
            byte[] dest = new byte[4];
            int n = XlsxWriter.ColumnLetter(col0, dest);
            var actual = System.Text.Encoding.ASCII.GetString(dest, 0, n);
            actual.ShouldBe(expected, $"col0={col0}");
        }

        [Fact]
        public void ColumnLetter_RandomColumns_MatchReference()
        {
            string Ref(int col0)
            {
                var s = "";
                int n = col0;
                while (true)
                {
                    s = (char)('A' + n % 26) + s;
                    n = n / 26 - 1;
                    if (n < 0) break;
                }
                return s;
            }
            byte[] dest = new byte[4];
            var rnd = new Random(20240707);
            for (int i = 0; i < 200; i++)
            {
                int col0 = rnd.Next(0, 16384);
                int n = XlsxWriter.ColumnLetter(col0, dest);
                var actual = System.Text.Encoding.ASCII.GetString(dest, 0, n);
                actual.ShouldBe(Ref(col0), $"col0={col0}");
            }
        }


        [Fact]
        public async Task Guardrail_PlainDto_ProducesWellFormedXlsx()
        {
            var bytes = Xlsx.ToBytes(new[]
            {
                new OrderDto { OrderNo = "A1", Amount = 10m, CreatedAt = new DateTime(2024, 1, 1) },
                new OrderDto { OrderNo = "B2", Amount = 20m, CreatedAt = new DateTime(2024, 2, 2) },
            });
            XlsxIO_TestSupport.AssertWellFormedXlsx(bytes);
        }

        [Fact]
        public async Task Guardrail_WithStyles_ProducesWellFormedXlsx()
        {
            var bytes = Xlsx.ToBytes(new[]
            {
                new OrderWithFormatDto { Amount = 1234.5m, Date = new DateTime(2024, 3, 3), Note = "x" },
            });
            XlsxIO_TestSupport.AssertWellFormedXlsx(bytes);
            var styles = XlsxIO_TestSupport.ReadStyles(bytes);
            styles.Count.ShouldBeGreaterThan(0);
        }

        [Fact]
        public async Task Guardrail_Features_ProducesWellFormedXlsx()
        {
            var p = new ExportProfile<OrderDto>()
                .WithMergeCells("A1:B2")
                .WithAutoFilter("A1:C10")
                .WithHyperlink("A2", "https://example.com");
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            XlsxIO_TestSupport.AssertWellFormedXlsx(bytes);
        }

        [Fact]
        public async Task Guardrail_SharedStrings_ProducesWellFormedXlsx()
        {
            var p = new ExportProfile<StringDto>().WithAutoSst(false);
            var bytes = Xlsx.ToBytes(new[]
            {
                new StringDto { Name = "hello" },
                new StringDto { Name = "world" },
                new StringDto { Name = "hello" },
            }, p);
            XlsxIO_TestSupport.AssertWellFormedXlsx(bytes);
        }


        [Fact]
        public void Complete_CanBeDrivenExplicitly_BeforeDispose()
        {
            var plan = RowPlanBuilder.BuildTyped(new ExportProfile<ReadbackOrder>());
            using var ms = new MemoryStream();
            var w = new XlsxWriter(ms);
            w.AddSheet("S1");
            w.ResolveColumnStyles(plan.Columns);
            w.WriteSheetMeta(plan.Columns, freezeHeader: false);
            w.WriteHeader(plan.Columns);
            w.WriteRows(new[] { new ReadbackOrder { Name = "x", Qty = 1, Price = 1m } }, plan);
            w.Complete();
            w.Dispose();
            var bytes = ms.ToArray();
            XlsxIO_TestSupport.AssertWellFormedXlsx(bytes);
            var rows = XlsxIO_TestSupport.ReadSheet(bytes, 1);
            rows[1][0].ShouldBe("x");
        }

        [Fact]
        public void Complete_IsIdempotent_WhenCalledMultipleTimes()
        {
            using var ms = new MemoryStream();
            var w = new XlsxWriter(ms);
            w.AddSheet("S1");
            w.Complete();
            w.Complete();
            w.Dispose();
            var bytes = ms.ToArray();
            using var z = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
            z.GetEntry("xl/workbook.xml").ShouldNotBeNull();
        }

        [Fact]
        public void Complete_WithoutSheet_ThrowsInsteadOfWritingBrokenWorkbook()
        {
            using var ms = new MemoryStream();
            using var writer = new XlsxWriter(ms);

            Should.Throw<InvalidOperationException>(() => writer.Complete());
        }


        [Theory]
        [InlineData("'BadName")]
        [InlineData("BadName'")]
        [InlineData("History")]
        [InlineData("history")]
        [InlineData("HIStory")]
        public void AddSheet_InvalidSheetName_Throws(string badName)
        {
            using var ms = new MemoryStream();
            using var w = new XlsxWriter(ms);
            Should.Throw<ArgumentException>(() => w.AddSheet(badName));
        }

        [Fact]
        public void AddSheet_ValidSheetName_Accepted()
        {
            using var ms = new MemoryStream();
            using var w = new XlsxWriter(ms);
            Should.NotThrow(() => w.AddSheet("NormalSheet"));
            Should.NotThrow(() => w.AddSheet("My-Sheet_1"));
        }

        [Fact]
        public void AddSheet_TrimsName_AsDocumented()
        {
            using var ms = new MemoryStream();
            using var w = new XlsxWriter(ms);
            w.AddSheet("  Trimmed  ");
            w.SheetNames.ShouldContain("Trimmed");
        }


        [Fact]
        public async Task Guardrail_MultiSheet_ReadBackEachSheet()
        {
            var plan = RowPlanBuilder.BuildTyped(new ExportProfile<ReadbackOrder>());
            using var ms = new MemoryStream();
            using (var w = new XlsxWriter(ms))
            {
                w.AddSheet("Sheet1");
                w.ResolveColumnStyles(plan.Columns);
                w.WriteSheetMeta(plan.Columns, freezeHeader: false);
                w.WriteHeader(plan.Columns);
                w.WriteRows(new[] { new ReadbackOrder { Name = "a1", Qty = 1, Price = 1.1m } }, plan);
                w.AddSheet("Sheet2");
                w.ResolveColumnStyles(plan.Columns);
                w.WriteSheetMeta(plan.Columns, freezeHeader: false);
                w.WriteHeader(plan.Columns);
                w.WriteRows(new[] { new ReadbackOrder { Name = "b2", Qty = 2, Price = 2.2m } }, plan);
                w.AddSheet("Sheet3");
                w.ResolveColumnStyles(plan.Columns);
                w.WriteSheetMeta(plan.Columns, freezeHeader: false);
                w.WriteHeader(plan.Columns);
                w.WriteRows(new[] { new ReadbackOrder { Name = "c3", Qty = 3, Price = 3.3m } }, plan);
            }
            var bytes = ms.ToArray();
            XlsxIO_TestSupport.AssertWellFormedXlsx(bytes);
            var s1 = XlsxIO_TestSupport.ReadSheet(bytes, 1);
            var s2 = XlsxIO_TestSupport.ReadSheet(bytes, 2);
            var s3 = XlsxIO_TestSupport.ReadSheet(bytes, 3);
            s1[1][0].ShouldBe("a1");
            s2[1][0].ShouldBe("b2");
            s3[1][0].ShouldBe("c3");
        }

        public sealed class WideDto
        {
            public string C00 { get; set; } = "";
            public string C01 { get; set; } = "";
            public string C02 { get; set; } = "";
            public string C03 { get; set; } = "";
            public string C04 { get; set; } = "";
            public string C05 { get; set; } = "";
            public string C06 { get; set; } = "";
            public string C07 { get; set; } = "";
            public string C08 { get; set; } = "";
            public string C09 { get; set; } = "";
            public string C10 { get; set; } = "";
            public string C11 { get; set; } = "";
            public string C12 { get; set; } = "";
            public string C13 { get; set; } = "";
            public string C14 { get; set; } = "";
            public string C15 { get; set; } = "";
            public string C16 { get; set; } = "";
            public string C17 { get; set; } = "";
            public string C18 { get; set; } = "";
            public string C19 { get; set; } = "";
            public string C20 { get; set; } = "";
            public string C21 { get; set; } = "";
            public string C22 { get; set; } = "";
            public string C23 { get; set; } = "";
            public string C24 { get; set; } = "";
            public string C25 { get; set; } = "";
            public string C26 { get; set; } = "";
            public string C27 { get; set; } = "";
            public string C28 { get; set; } = "";
            public string C29 { get; set; } = "";
        }

        [Fact]
        public async Task Guardrail_WideTable_Beyond26Columns_ReadBack()
        {
            var bytes = Xlsx.ToBytes(new[] { new WideDto { C00 = "x", C26 = "y", C29 = "z" } });
            XlsxIO_TestSupport.AssertWellFormedXlsx(bytes);
            var rows = XlsxIO_TestSupport.ReadSheet(bytes, 1);
            rows.Count.ShouldBeGreaterThanOrEqualTo(2);
            var header = string.Join(",", rows[0].Where(x => x is not null).Select(x => x!));
            header.ShouldContain("C26");
            header.ShouldContain("C29");
            rows[1][0].ShouldBe("x");
        }
    }
}
