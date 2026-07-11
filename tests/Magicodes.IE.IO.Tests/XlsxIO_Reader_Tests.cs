
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.IO.Compression;
using Magicodes.IE.IO;
using MiniExcelLibs;
using Shouldly;
using Xunit;

namespace Magicodes.IE.IO.Tests
{
    public partial class XlsxIO_Tests
    {
        [Fact]
        public async Task XlsxRead_ReadsInlineStringsAndNumbers()
        {
            var bytes = Xlsx.ToBytes(new[]
            {
                new OrderDto { OrderNo = "A", Amount = 10m, CreatedAt = new DateTime(2024, 1, 1) },
                new OrderDto { OrderNo = "B", Amount = 20m, CreatedAt = new DateTime(2024, 1, 2) },
            });
            var list = Xlsx.Read<OrderDto>(new MemoryStream(bytes)).ToList();
            list.Count.ShouldBe(2);
            list[0].OrderNo.ShouldBe("A");
            list[0].Amount.ShouldBe(10m);
            list[1].OrderNo.ShouldBe("B");
        }

        [Fact]
        public async Task XlsxRead_RoundTrip_PreservesData()
        {
            var src = new[]
            {
                new OrderDto { OrderNo = "X", Amount = 1.5m, CreatedAt = new DateTime(2024, 5, 1) },
                new OrderDto { OrderNo = "Y", Amount = 2.5m, CreatedAt = new DateTime(2024, 5, 2) },
            };
            var bytes = Xlsx.ToBytes(src);
            var dst = Xlsx.Read<OrderDto>(new MemoryStream(bytes)).ToList();
            dst.Count.ShouldBe(src.Length);
            for (int i = 0; i < src.Length; i++)
            {
                dst[i].OrderNo.ShouldBe(src[i].OrderNo);
                dst[i].Amount.ShouldBe(src[i].Amount);
                dst[i].CreatedAt.ShouldBe(src[i].CreatedAt);
            }
        }

        [Fact]
        public async Task XlsxRead_AsyncStreaming()
        {
            var bytes = Xlsx.ToBytes(Enumerable.Range(0, 20).Select(i => new OrderDto { OrderNo = $"R{i}" }));
            var list = new List<OrderDto>();
            await foreach (var o in Xlsx.ReadAsync<OrderDto>(new MemoryStream(bytes)))
                list.Add(o);
            list.Count.ShouldBe(20);
        }

        [Fact]
        public void XlsxReader_EmptyRowDoesNotConsumeFollowingRow()
        {
            using var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                var sheet = zip.CreateEntry("xl/worksheets/sheet1.xml");
                using var es = sheet.Open();
                using var sw = new StreamWriter(es);
                sw.Write("<?xml version=\"1.0\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
                    + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>OrderNo</t></is></c></row>"
                    + "<row r=\"2\"/><row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is><t>after</t></is></c></row>"
                    + "</sheetData></worksheet>");
            }

            ms.Position = 0;
            using var reader = new XlsxReader(ms);
            reader.ReadHeader().ShouldBe(new[] { "OrderNo" });
            reader.ReadNextRowView().ShouldBeEmpty();
            reader.ReadNextRowView()![0].ShouldBe("after");
        }

        [Fact]
        public void XlsxReader_OutOfOrderCellRefs_PlacedByColumnIndex()
        {
            using var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                var sheet = zip.CreateEntry("xl/worksheets/sheet1.xml");
                using var es = sheet.Open();
                using var sw = new StreamWriter(es);
                sw.Write("<?xml version=\"1.0\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
                    + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>H1</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>H2</t></is></c><c r=\"C1\" t=\"inlineStr\"><is><t>H3</t></is></c></row>"
                    + "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>a</t></is></c><c r=\"C2\" t=\"inlineStr\"><is><t>c</t></is></c><c r=\"B2\" t=\"inlineStr\"><is><t>b</t></is></c></row>"
                    + "</sheetData></worksheet>");
            }

            ms.Position = 0;
            using var reader = new XlsxReader(ms);
            reader.ReadHeader().ShouldBe(new[] { "H1", "H2", "H3" });
            // Cells emitted out of order (A, C, B) must still land at their column index,
            // not be appended in document order.
            reader.ReadNextRowView()!.ShouldBe(new[] { "a", "b", "c" });
        }

        [Fact]
        public void XlsxReader_InlineStringRichText_ConcatenatesRuns()
        {
            using var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                var sheet = zip.CreateEntry("xl/worksheets/sheet1.xml");
                using var es = sheet.Open();
                using var sw = new StreamWriter(es);
                sw.Write("<?xml version=\"1.0\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
                    + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>H</t></is></c></row>"
                    + "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><r><t>foo</t></r><r><t>bar</t></r></is></c></row>"
                    + "</sheetData></worksheet>");
            }

            ms.Position = 0;
            using var reader = new XlsxReader(ms);
            reader.ReadHeader().ShouldBe(new[] { "H" });
            // Multiple <t> runs (rich text) must be concatenated, exercising the StringBuilder path.
            reader.ReadNextRowView()![0].ShouldBe("foobar");
        }

        [Fact]
        public void XlsxRead_ULongAboveLongMaxValue_DoesNotWrapNegative()
        {
            // ulong > long.MaxValue used to be cast to (long) in the generated getter and
            // silently wrap to a negative value; it must now survive round-trip as positive.
            // 1 << 63 is exactly representable as double, so it round-trips within Excel's
            // IEEE754-double precision.
            var src = new[]
            {
                new ULongDto { Name = "big", Value = 1UL << 63 },
            };
            var bytes = Xlsx.ToBytes(src);
            var dst = Xlsx.Read<ULongDto>(new MemoryStream(bytes)).Single();
            dst.Name.ShouldBe("big");
            dst.Value.ShouldBe(1UL << 63);
            dst.Value.ShouldBeGreaterThan((ulong)long.MaxValue);
        }

        [Fact]
        public void XlsxWrite_ULongBackedEnumAboveLongMax_DoesNotOverflow()
        {
            // A ulong-backed enum with a value > long.MaxValue used to be cast to (long) and
            // silently wrap to negative (cell-writer path) or throw OverflowException
            // (typed-getter / ToCellValue paths). It must now write a positive number.
            var bytes = Xlsx.ToBytes(new[] { new ULongEnumDto { Name = "x", Flag = BigFlag.Huge } });
            // Read the same cell back as double (Enum.TryParse rejects the scientific-notation
            // text, so read as the underlying numeric to verify the written value).
            var asDouble = Xlsx.Read<ULongEnumAsDoubleDto>(new MemoryStream(bytes)).Single();
            asDouble.Name.ShouldBe("x");
            // Before the fix this was -9.223372036854776E+18 (negative wrap). The written cell
            // must be the positive 2^63 (= (double)(1UL << 63)).
            asDouble.Flag.ShouldBe((double)(1UL << 63), $"cell was {asDouble.Flag}");
            asDouble.Flag.ShouldBeGreaterThan(0);
        }

        [Fact]
        public void XlsxWrite_ColumnsWithoutIndex_PreserveDeclarationOrder()
        {
            // List.Sort is not stable; columns sharing the default Index (int.MaxValue) must
            // still come out in declaration order via the tiebreaker, deterministically.
            var bytes = Xlsx.ToBytes(new[] { new DeclOrderDto { Alpha = "a", Bravo = "b", Charlie = "c" } });
            using var ms = new MemoryStream(bytes);
            using var reader = new XlsxReader(ms);
            reader.ReadHeader().ShouldBe(new[] { "Alpha", "Bravo", "Charlie" });
        }

        [Fact]
        public void XlsxRead_StructPropertiesAreAssigned()
        {
            var bytes = Xlsx.ToBytes(new[] { new TestOrderRecord("SO-1", 12.5m) });
            var result = Xlsx.Read<TestOrderRecord>(new MemoryStream(bytes)).Single();

            result.OrderNo.ShouldBe("SO-1");
            result.Amount.ShouldBe(12.5m);
        }

        [Fact]
        public void CellValueParser_Parses1904DatesAndDecimalExponents()
        {
            CellValueParser.TryParse("0", typeof(DateTime), out var date, date1904: true).ShouldBeTrue();
            ((DateTime)date!).ShouldBe(new DateTime(1904, 1, 1));

            CellValueParser.TryParse("1E+2", typeof(decimal), out var number).ShouldBeTrue();
            ((decimal)number!).ShouldBe(100m);
        }

        [Fact]
        public async Task ReadAsync_CancellationRequested_ThrowsOperationCanceled()
        {
            var bytes = Xlsx.ToBytes(Enumerable.Range(0, 100).Select(i => new OrderDto { OrderNo = $"R{i}" }));
            using var ms = new MemoryStream(bytes);
            var cts = new CancellationTokenSource();
            cts.Cancel();
            var list = new List<OrderDto>();
            await Should.ThrowAsync<OperationCanceledException>(async () =>
            {
                await foreach (var o in Xlsx.ReadAsync<OrderDto>(ms, cancellationToken: cts.Token))
                    list.Add(o);
            });
            list.Count.ShouldBe(0);
        }

        [Fact]
        public async Task ReadAsync_CancellationRequested_BeforeStart_DoesNotTouchStream()
        {
            var bytes = Xlsx.ToBytes(Enumerable.Range(0, 5).Select(i => new OrderDto { OrderNo = $"R{i}" }));
            using var tracking = new XlsxIO_TestSupport.TrackingReadStream(new MemoryStream(bytes));
            var cts = new CancellationTokenSource();
            cts.Cancel();

            await Should.ThrowAsync<OperationCanceledException>(async () =>
            {
                await foreach (var _ in Xlsx.ReadAsync<OrderDto>(tracking, cancellationToken: cts.Token))
                {
                }
            });

            tracking.ReadCount.ShouldBe(0);
            tracking.ReadAsyncCount.ShouldBe(0);
            tracking.SeekCount.ShouldBe(0);
            tracking.PositionGetCount.ShouldBe(0);
            tracking.PositionSetCount.ShouldBe(0);
            tracking.LengthGetCount.ShouldBe(0);
        }

        [Fact]
        public async Task XlsxRead_RoundTrip_DirectSurface()
        {
            var bytes = Xlsx.ToBytes(new[] { new OrderDto { OrderNo = "Z" } });
            var list = Xlsx.Read<OrderDto>(new MemoryStream(bytes)).ToList();
            list.Count.ShouldBe(1);
            list[0].OrderNo.ShouldBe("Z");
        }

        [Fact]
        public void ImportProfile_ExplicitColumnMapping_WinsOverHeaderName()
        {
            using var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                var sheet = zip.CreateEntry("xl/worksheets/sheet1.xml");
                using var es = sheet.Open();
                using var sw = new StreamWriter(es);
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                        + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                        + "<sheetData>"
                        + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Amount</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>Other</t></is></c></row>"
                        + "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>X</t></is></c><c r=\"B2\"><v>12</v></c></row>"
                        + "</sheetData></worksheet>");
            }

            ms.Position = 0;
            var profile = new XlsxReadOptions<OrderDto>()
                .MapColumn(0, nameof(OrderDto.OrderNo))
                .MapColumn(1, nameof(OrderDto.Amount));

            var list = Xlsx.Read<OrderDto>(ms, profile).ToList();
            list.Count.ShouldBe(1);
            list[0].OrderNo.ShouldBe("X");
            list[0].Amount.ShouldBe(12m);
        }

        [Fact]
        public void ImportProfile_HeaderMapping_MapsByHeaderText()
        {
            using var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                var sheet = zip.CreateEntry("xl/worksheets/sheet1.xml");
                using var es = sheet.Open();
                using var sw = new StreamWriter(es);
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                        + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                        + "<sheetData>"
                        + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>订单号</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>金额</t></is></c></row>"
                        + "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>R1</t></is></c><c r=\"B2\"><v>12</v></c></row>"
                        + "</sheetData></worksheet>");
            }

            ms.Position = 0;
            var profile = new XlsxReadOptions<OrderDto>()
                .MapHeader("订单号", nameof(OrderDto.OrderNo))
                .MapHeader("金额", nameof(OrderDto.Amount));

            var list = Xlsx.Read<OrderDto>(ms, profile).ToList();
            list.Count.ShouldBe(1);
            list[0].OrderNo.ShouldBe("R1");
            list[0].Amount.ShouldBe(12m);
        }

        [Fact]
        public void XlsxReader_UsesWorkbookRelationshipOrder_NotHardcodedSheet1Xml()
        {
            using var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                var workbook = zip.CreateEntry("xl/workbook.xml");
                using (var es = workbook.Open())
                using (var sw = new StreamWriter(es))
                {
                    sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                        + "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                        + "<sheets>"
                        + "<sheet name=\"Preferred\" sheetId=\"1\" r:id=\"rIdPreferred\"/>"
                        + "<sheet name=\"Ignored\" sheetId=\"2\" r:id=\"rIdIgnored\"/>"
                        + "</sheets></workbook>");
                }

                var workbookRels = zip.CreateEntry("xl/_rels/workbook.xml.rels");
                using (var es = workbookRels.Open())
                using (var sw = new StreamWriter(es))
                {
                    sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                        + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                        + "<Relationship Id=\"rIdPreferred\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet2.xml\"/>"
                        + "<Relationship Id=\"rIdIgnored\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
                        + "</Relationships>");
                }

                var sheet1 = zip.CreateEntry("xl/worksheets/sheet1.xml");
                using (var es = sheet1.Open())
                using (var sw = new StreamWriter(es))
                {
                    sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                        + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
                        + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>OrderNo</t></is></c></row>"
                        + "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>WRONG</t></is></c></row>"
                        + "</sheetData></worksheet>");
                }

                var sheet2 = zip.CreateEntry("xl/worksheets/sheet2.xml");
                using (var es = sheet2.Open())
                using (var sw = new StreamWriter(es))
                {
                    sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                        + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData>"
                        + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>OrderNo</t></is></c></row>"
                        + "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>RIGHT</t></is></c></row>"
                        + "</sheetData></worksheet>");
                }
            }

            ms.Position = 0;
            var list = Xlsx.Read<OrderDto>(ms).ToList();
            list.Count.ShouldBe(1);
            list[0].OrderNo.ShouldBe("RIGHT");
        }

        [Fact]
        public async Task Read_XlsxWrittenByMiniExcel_RoundTrips()
        {
            var path = Path.Combine(Path.GetTempPath(), $"miniexcel_{Guid.NewGuid():N}.xlsx");
            try
            {
                MiniExcel.SaveAs(path, new[]
                {
                    new { OrderNo = "M1", Amount = 10 },
                    new { OrderNo = "M2", Amount = 20 },
                });
                var items = Xlsx.Read<OrderDto>(File.OpenRead(path)).ToList();
                items.Count.ShouldBe(2);
                items[0].OrderNo.ShouldBe("M1");
            }
            finally
            {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public async Task Read_SharedStrings_RealXlsxFormat()
        {
            using var ms = new MemoryStream();
            using (var writer = new XlsxWriter(ms))
            {
                writer.AddSheet("S1");
                writer.EnableSharedStrings();
                var cols = new[] { new ColumnMeta("Name", "Name", null, null, false, 0, 0) };
                writer.ResolveColumnStyles(cols);
                writer.WriteSheetMeta(cols, freezeHeader: true);
                writer.WriteHeader(cols);
                var getters = new Func<StringDto, CellValue>[] { o => CellValue.FromString(o.Name) };
                var plan = new TypedRowPlan<StringDto>(cols, new Func<object?, CellValue>[0], getters, new int[1], new Action<XlsxWriter.XlsxRowWriter, StringDto, int>?[1], hasFormulas: false);
                writer.WriteRows(new[]
                {
                    new StringDto { Name = "x" }, new StringDto { Name = "x" }, new StringDto { Name = "y" }
                }, plan);
            }
            var list = Xlsx.Read<StringDto>(new MemoryStream(ms.ToArray())).ToList();
            list.Count.ShouldBe(3);
            list[0].Name.ShouldBe("x");
            list[2].Name.ShouldBe("y");
        }

        [Fact]
        public async Task RoundTrip_SelfWrittenXlsx_CanBeParsed()
        {
            var bytes = Xlsx.ToBytes(Enumerable.Range(0, 50).Select(i => new OrderDto { OrderNo = $"O{i}" }));
            using var reader = new XlsxReader(new MemoryStream(bytes));
            var headers = reader.ReadHeader();
            headers.ShouldContain("OrderNo");
            int rowCount = 0;
            while (reader.ReadNextRow() is not null) rowCount++;
            rowCount.ShouldBe(50);
        }

        [Fact]
        public async Task ReadSharedStrings_Handled()
        {
            using var ms = new MemoryStream();
            using (var writer = new XlsxWriter(ms))
            {
                writer.AddSheet("S1");
                writer.EnableSharedStrings();
                var cols = new[] { new ColumnMeta("Name", "Name", null, null, false, 0, 0) };
                writer.ResolveColumnStyles(cols);
                writer.WriteSheetMeta(cols, freezeHeader: true);
                writer.WriteHeader(cols);
                var getters = new Func<StringDto, CellValue>[] { o => CellValue.FromString(o.Name) };
                var plan = new TypedRowPlan<StringDto>(cols, new Func<object?, CellValue>[0], getters, new int[1], new Action<XlsxWriter.XlsxRowWriter, StringDto, int>?[1], hasFormulas: false);
                writer.WriteRows(Enumerable.Range(0, 1000).Select(_ => new StringDto { Name = "same" }), plan);
            }
            var list = Xlsx.Read<StringDto>(new MemoryStream(ms.ToArray())).ToList();
            list.Count.ShouldBe(1000);
            list[500].Name.ShouldBe("same");
        }

        [Fact]
        public async Task Read_InlineStringRichText_AggregatesAllRuns()
        {
            using var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                var sheet = zip.CreateEntry("xl/worksheets/sheet1.xml");
                using var es = sheet.Open();
                using var sw = new StreamWriter(es);
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                    + "<sheetData>"
                    + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Name</t></is></c></row>"
                    + "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><r><t>Hel</t></r><r><t>lo</t></r></is></c></row>"
                    + "</sheetData></worksheet>");
            }

            ms.Position = 0;
            var list = Xlsx.Read<StringDto>(ms).ToList();
            list.Count.ShouldBe(1);
            list[0].Name.ShouldBe("Hello");
        }

        [Fact]
        public async Task Read_DateCell_TypedDateValue_RoundTrips()
        {
            using var ms = new MemoryStream();
            using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
            {
                var sheet = zip.CreateEntry("xl/worksheets/sheet1.xml");
                using var es = sheet.Open();
                using var sw = new StreamWriter(es);
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                    + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                    + "<sheetData>"
                    + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>CreatedAt</t></is></c></row>"
                    + "<row r=\"2\"><c r=\"A2\" t=\"d\"><v>2024-01-01T00:00:00Z</v></c></row>"
                    + "</sheetData></worksheet>");
            }

            ms.Position = 0;
            var list = Xlsx.Read<OrderDto>(ms).ToList();
            list.Count.ShouldBe(1);
            list[0].CreatedAt.ShouldBe(new DateTime(2024, 1, 1, 0, 0, 0, DateTimeKind.Utc));
        }

        [Fact]
        public async Task Query_OnParseError_InvokedOnInvalidNumber()
        {
            using var ms = new MemoryStream();
            using (var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
            {
                var entry = zip.CreateEntry("xl/worksheets/sheet1.xml");
                using var es = entry.Open();
                using var sw = new StreamWriter(es);
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                  + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                  + "<sheetData>"
                  + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>OrderNo</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>Amount</t></is></c></row>"
                  + "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>ON</t></is></c><c r=\"B2\"><v>not-a-number</v></c></row>"
                  + "</sheetData></worksheet>");
            }
            ms.Position = 0;
            var errors = new List<XlsxReadErrorInfo>();
            var list = Xlsx.Read<OrderDto>(ms, onParseError: e => errors.Add(e)).ToList();
            list.Count.ShouldBe(1);
            errors.Count.ShouldBe(1);
            errors[0].PropertyName.ShouldBe("Amount");
            errors[0].RawCellValue.ShouldBe("not-a-number");
            errors[0].TargetTypeName.ShouldBe("Decimal");
            errors[0].RowIndex.ShouldBe(0);
            errors[0].ColIndex.ShouldBe(1);
            errors[0].Exception.ShouldNotBeNull();
        }

        [Fact]
        public async Task Query_OnParseError_InvokedOnInvalidBool()
        {
            using var ms = new MemoryStream();
            using (var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
            {
                var entry = zip.CreateEntry("xl/worksheets/sheet1.xml");
                using var es = entry.Open();
                using var sw = new StreamWriter(es);
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                  + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                  + "<sheetData>"
                  + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>Enabled</t></is></c></row>"
                  + "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>A</t></is></c><c r=\"B2\" t=\"inlineStr\"><is><t>maybe</t></is></c></row>"
                  + "</sheetData></worksheet>");
            }
            ms.Position = 0;
            var errors = new List<XlsxReadErrorInfo>();
            var list = Xlsx.Read<BoolDto>(ms, onParseError: e => errors.Add(e)).ToList();
            list.Count.ShouldBe(1);
            errors.Count.ShouldBe(1);
            errors[0].PropertyName.ShouldBe("Enabled");
            errors[0].RawCellValue.ShouldBe("maybe");
            errors[0].TargetTypeName.ShouldBe("Boolean");
        }

        [Fact]
        public void Query_OnParseError_InvokedOnInvalidTypedBooleanCell()
        {
            using var ms = new MemoryStream();
            using (var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
            {
                var entry = zip.CreateEntry("xl/worksheets/sheet1.xml");
                using var es = entry.Open();
                using var sw = new StreamWriter(es);
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                  + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                  + "<sheetData>"
                  + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Name</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>Enabled</t></is></c></row>"
                  + "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>A</t></is></c><c r=\"B2\" t=\"b\"><v>2</v></c></row>"
                  + "</sheetData></worksheet>");
            }
            ms.Position = 0;
            var errors = new List<XlsxReadErrorInfo>();
            var list = Xlsx.Read<BoolDto>(ms, onParseError: e => errors.Add(e)).ToList();
            list.Count.ShouldBe(1);
            list[0].Enabled.ShouldBeFalse();
            errors.Count.ShouldBe(1);
            errors[0].PropertyName.ShouldBe("Enabled");
            errors[0].RawCellValue.ShouldBe("2");
            errors[0].TargetTypeName.ShouldBe("Boolean");
        }

        [Fact]
        public void ReadNextRow_HandlesSparseColumns()
        {
            using var ms = new MemoryStream();
            using (var zip = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Create, leaveOpen: true))
            {
                var entry = zip.CreateEntry("xl/worksheets/sheet1.xml");
                using var es = entry.Open();
                using var sw = new StreamWriter(es);
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                  + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                  + "<sheetData>"
                  + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>A</t></is></c><c r=\"B1\" t=\"inlineStr\"><is><t>B</t></is></c><c r=\"C1\" t=\"inlineStr\"><is><t>C</t></is></c></row>"
                  + "<row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>ON</t></is></c><c r=\"C2\" t=\"inlineStr\"><is><t>VAL</t></is></c></row>"
                  + "</sheetData></worksheet>");
            }
            ms.Position = 0;
            using var reader = new XlsxReader(ms);
            var headers = reader.ReadHeader();
            headers.ShouldBe(new[] { "A", "B", "C" });
            var row = reader.ReadNextRow();
            row.ShouldNotBeNull();
            row.Count.ShouldBe(3);
            row[0].ShouldBe("ON");
            row[1].ShouldBeNull();
            row[2].ShouldBe("VAL");
        }



        private readonly struct Money
        {
            public readonly decimal Value;
            public Money(decimal v) { Value = v; }
            public override string ToString() => Value.ToString("F2");
        }


        private sealed class MoneyConverter : CellConverter<Money>
        {
            public override bool Read(string cell, out Money value)
            {
                if (decimal.TryParse(cell, System.Globalization.NumberStyles.Number,
                    System.Globalization.CultureInfo.InvariantCulture, out var d))
                {
                    value = new Money(d);
                    return true;
                }
                value = default;
                return false;
            }
        }

        private sealed class MoneyDto
        {
            public string? Label { get; set; }
            public Money Amount { get; set; }
        }

        [Fact]
        public void CellConverter_Read_CustomType_ParsesCorrectly()
        {
            var bytes = Xlsx.ToBytes(new[]
            {
                new MoneyDto { Label = "A", Amount = new Money(10.5m) },
                new MoneyDto { Label = "B", Amount = new Money(20.0m) },
            });

            var opts = new XlsxReadOptions<MoneyDto>();
            opts.WithConverter(new MoneyConverter());

            var list = Xlsx.Read<MoneyDto>(new MemoryStream(bytes), opts).ToList();
            list.Count.ShouldBe(2);
            list[0].Label.ShouldBe("A");
            list[0].Amount.Value.ShouldBe(10.5m);
            list[1].Label.ShouldBe("B");
            list[1].Amount.Value.ShouldBe(20.0m);
        }

        [Fact]
        public void ExportProfile_Freeze_IsPublic_And_ThrowsOnMutation()
        {
            var profile = new ExportProfile<OrderDto>();
            profile.Sheet("test");
            profile.IsFrozen.ShouldBeFalse();
            profile.Freeze();
            profile.IsFrozen.ShouldBeTrue();
            Should.Throw<InvalidOperationException>(() => profile.Sheet("other"));
        }

        [Fact]
        public void XlsxException_ContainsStructuredLocation()
        {
            var ex = new XlsxException("bad cell", rowIndex: 5, colIndex: 2, cellRef: "C6");
            ex.RowIndex.ShouldBe(5);
            ex.ColIndex.ShouldBe(2);
            ex.CellRef.ShouldBe("C6");
            ex.Message.ShouldContain("Row=5");
            ex.Message.ShouldContain("Col=2");
            ex.Message.ShouldContain("CellRef=C6");
        }

        [Fact]
        public async Task WriteAsync_WithStrictCellReferences_PassesThrough()
        {
            var data = new[]
            {
                new OrderDto { OrderNo = "X", Amount = 1m, CreatedAt = new DateTime(2024, 1, 1) },
            };

            using var ms = new MemoryStream();
            await Xlsx.WriteAsync(ms, DataToAsync(data),
                options: new XlsxWriteOptions { StrictCellReferences = false, Compression = CompressionLevel.NoCompression });
            ms.Length.ShouldBeGreaterThan(0);

            ms.Position = 0;
            using var zip = new System.IO.Compression.ZipArchive(ms, ZipArchiveMode.Read);
            var sheetEntry = zip.GetEntry("xl/worksheets/sheet1.xml");
            sheetEntry.ShouldNotBeNull();
            using var es = sheetEntry!.Open();
            var content = new StreamReader(es).ReadToEnd();
            content.ShouldContain("t=\"n\"");
            content.ShouldNotContain("r=\"A2\"");
        }

        [Fact]
        public async Task WriteAsync_DelayedAsyncEnumerable_WritesAllRowsBeforeDisposingWriter()
        {
            static async IAsyncEnumerable<OrderDto> DelayedData()
            {
                await Task.Delay(10);
                yield return new OrderDto { OrderNo = "A" };
                await Task.Delay(10);
                yield return new OrderDto { OrderNo = "B" };
            }

            using var ms = new MemoryStream();
            await Xlsx.WriteAsync(ms, DelayedData());

            var rows = XlsxIO_TestSupport.ReadSheet(ms.ToArray());
            rows.Count.ShouldBe(3);
            rows[1][0].ShouldBe("A");
            rows[2][0].ShouldBe("B");
        }

        private static async IAsyncEnumerable<T> DataToAsync<T>(IEnumerable<T> data, [System.Runtime.CompilerServices.EnumeratorCancellation] CancellationToken ct = default)
        {
            foreach (var item in data)
            {
                yield return item;
                await Task.Yield();
            }
        }
    }
}
