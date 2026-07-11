
using System;
using System.Buffers;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using Magicodes.IE.IO;
using Shouldly;
using Xunit;

namespace Magicodes.IE.IO.Tests
{
    public partial class XlsxIO_Tests
    {
        [Fact]
        public async Task SaveAsync_PlainDto_WritesFile()
        {
            var path = Path.Combine(Path.GetTempPath(), $"xlsx_io_test_{Guid.NewGuid():N}.xlsx");
            try
            {
                Xlsx.Write(path, new[] { new OrderDto { OrderNo = "A1", Amount = 10m, CreatedAt = new DateTime(2024, 1, 1) } });
                File.Exists(path).ShouldBeTrue();
                var bytes = await File.ReadAllBytesAsync(path);
                var rows = XlsxIO_TestSupport.ReadSheet(bytes);
                rows.Count.ShouldBe(2);
                rows[1][0].ShouldBe("A1");
            }
            finally
            {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public async Task SaveAsBytes_PlainDto_HeaderIsPropertyName()
        {
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() });
            var rows = XlsxIO_TestSupport.ReadSheet(bytes);
            rows[0][0].ShouldBe("OrderNo");
            rows[0][1].ShouldBe("Amount");
            rows[0][2].ShouldBe("CreatedAt");
        }

        [Fact]
        public async Task WriteAsync_BufferWriter_PlainDto_WritesValidXlsx()
        {
            var output = new ArrayBufferWriter<byte>();
            Xlsx.Write(output, new[] { new OrderDto { OrderNo = "BufferWriter" } });
            var rows = XlsxIO_TestSupport.ReadSheet(output.WrittenMemory.ToArray());
            rows[1][0].ShouldBe("BufferWriter");
        }

        [Fact]
        public async Task WriteAsync_Stream_WithOptions_WritesValidXlsx()
        {
            using var ms = new MemoryStream();

            Xlsx.Write(ms, new[] { new OrderDto { OrderNo = "W1" } }, p => p
                .Sheet("Orders")
                .Column(x => x.OrderNo, c => c.WithName("订单号")));

            var rows = XlsxIO_TestSupport.ReadSheet(ms.ToArray());
            rows[0][0].ShouldBe("订单号");
            rows[1][0].ShouldBe("W1");
        }

        [Fact]
        public async Task WriteAsync_ProfileSheetName_WritesGivenSheet()
        {
            using var ms = new MemoryStream();

            Xlsx.Write(ms, new[] { new OrderDto { OrderNo = "P1" } }, p => p
                .Sheet("ProfileSheet")
                .Column(x => x.OrderNo, c => c.WithName("订单号")));

            var workbookXml = XlsxIO_TestSupport.ReadEntry(ms.ToArray(), "xl/workbook.xml");
            workbookXml.ShouldContain("ProfileSheet");
        }

        [Fact]
        public async Task WriteAsync_BufferWriter_WithOptions_WritesValidXlsx()
        {
            var output = new ArrayBufferWriter<byte>();

            Xlsx.Write(output, new[] { new OrderDto { OrderNo = "BW1" } }, p => p
                .Column(x => x.OrderNo, c => c.WithName("订单号")));

            var rows = XlsxIO_TestSupport.ReadSheet(output.WrittenMemory.ToArray());
            rows[0][0].ShouldBe("订单号");
            rows[1][0].ShouldBe("BW1");
        }

        [Fact]
        public async Task SaveAs_WithProfileInline_AppliesHeaderAndIgnore()
        {
            var profile = new ExportProfile<OrderWithAttributesDto>()
                .Column(x => x.OrderNo, c => c.WithName("订单号"))
                .Ignore(x => x.Secret);
            var bytes = Xlsx.ToBytes(new[] { new OrderWithAttributesDto() }, profile);
            var rows = XlsxIO_TestSupport.ReadSheet(bytes);
            rows[0].ShouldContain("订单号");
            rows[0].ShouldNotContain("Secret");
        }

        [Fact]
        public async Task SaveAs_WithProfileClass_AppliesConfig()
        {
            var p = new ExportProfile<OrderWithAttributesDto>()
                .Column(x => x.OrderNo, c => c.WithName("订单号"));
            var bytes = Xlsx.ToBytes(new[] { new OrderWithAttributesDto() }, p);
            var rows = XlsxIO_TestSupport.ReadSheet(bytes);
            rows[0].ShouldContain("订单号");
        }

        [Fact]
        public async Task SaveAs_WithHeader_RenamesColumn()
        {
            var p = new ExportProfile<OrderDto>().Column(x => x.OrderNo, c => c.WithName("Order ID"));
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var rows = XlsxIO_TestSupport.ReadSheet(bytes);
            rows[0][0].ShouldBe("Order ID");
        }

        [Fact]
        public async Task SaveAs_WithHiddenColumn_KeepsDataAlignedAndWritesHiddenFlag()
        {
            var p = new ExportProfile<OrderDto>()
                .Column(x => x.OrderNo, c => c.WithName("订单号"))
                .Column(x => x.Amount, c => c.WithHidden())
                .Column(x => x.CreatedAt, c => c.WithName("创建时间"));

            var bytes = Xlsx.ToBytes(new[]
            {
                new OrderDto { OrderNo = "A1", Amount = 12.34m, CreatedAt = new DateTime(2024, 1, 1) }
            }, p);

            var rows = XlsxIO_TestSupport.ReadSheet(bytes);
            rows[0].Count.ShouldBe(3);
            rows[1].Count.ShouldBe(3);
            rows[0][0].ShouldBe("订单号");
            rows[0][1].ShouldBe("Amount");
            rows[0][2].ShouldBe("创建时间");
            rows[1][1].ShouldBe("12.34");

            var sheetXml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            sheetXml.ShouldContain("hidden=\"1\"");
            sheetXml.ShouldContain("r=\"B1\"");
            sheetXml.ShouldContain("r=\"B2\"");
        }

        [Fact]
        public void SourceGenerator_RegistersTypedGetters_ForExportableDto()
        {
            var getters = XlsxGeneratedTypedGettersRegistry.TryGet(typeof(ExportableDto));
            getters.ShouldNotBeNull();
            getters!.Keys.ShouldContain("OrderNo");
            getters.Keys.ShouldContain("Amount");

        }

        [Fact]
        public void SourceGenerator_RegistersReaderMetadata_ForExportableDto()
        {
            var metadata = XlsxGeneratedTypeMetadataRegistry.TryGet<ExportableDto>();
            metadata.ShouldNotBeNull();
            metadata!.Count.ShouldBe(2);
            metadata.Single(x => x.Name == "OrderNo").Setter.ShouldNotBeNull();
            metadata.Single(x => x.Name == "Amount").Setter.ShouldNotBeNull();

            var bytes = Xlsx.ToBytes(new[]
            {
                new ExportableDto { OrderNo = "AOT", Amount = 12.5m }
            });

            using var stream = new MemoryStream(bytes);
            var row = Xlsx.Read<ExportableDto>(stream).Single();
            row.OrderNo.ShouldBe("AOT");
            row.Amount.ShouldBe(12.5m);
        }

        [Fact]
        public async Task SaveAs_SourceGeneratedRowWriter_RespectsIgnoreAndIndex()
        {
            var p = new ExportProfile<ExportableReorderedDto>()
                .Ignore(x => x.B)
                .Column(x => x.C, c => c.WithIndex(0))
                .Column(x => x.A, c => c.WithIndex(1));

            var bytes = Xlsx.ToBytes(new[]
            {
                new ExportableReorderedDto { A = "A1", B = "B1", C = "C1" }
            }, p);

            var rows = XlsxIO_TestSupport.ReadSheet(bytes);
            rows[0].ShouldBe(new List<string?> { "C", "A" });
            rows[1].ShouldBe(new List<string?> { "C1", "A1" });

        }

        [Fact]
        public void Profile_Freeze_ThrowsOnFurtherMutation()
        {
            var p = new ExportProfile<OrderDto>();
            p.Column(x => x.OrderNo, c => c.WithName("X"));
            _ = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            Should.Throw<InvalidOperationException>(() => p.Column(x => x.OrderNo, c => c.WithName("Y")));
        }

        [Fact]
        public async Task SaveAs_EmptyData_WritesHeaderOnly()
        {
            var bytes = Xlsx.ToBytes(Array.Empty<OrderDto>());
            var rows = XlsxIO_TestSupport.ReadSheet(bytes);
            rows.Count.ShouldBe(1);
        }

        [Fact]
        public async Task SaveAs_LargeData_WritesAllRows()
        {
            var data = Enumerable.Range(0, 1000).Select(i => new OrderDto { OrderNo = $"O{i}" }).ToArray();
            var bytes = Xlsx.ToBytes(data);
            var rows = XlsxIO_TestSupport.ReadSheet(bytes);
            rows.Count.ShouldBe(1001);
        }

        [Fact]
        public async Task SaveAs_Comments_InvalidXmlChars_AreSanitized()
        {
            using var ms = new MemoryStream();
            using (var writer = new XlsxWriter(ms))
            {
                writer.AddSheet("S1");
                writer.AddComment(new Comment(0, 0, "A\u0001B", "x\u0002y"));
            }

            var bytes = ms.ToArray();
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/comments1.xml");
            XlsxIO_TestSupport.AssertWellFormedXml(xml, "comments1.xml");
            xml.ShouldNotContain("\u0001");
            xml.ShouldNotContain("\u0002");
        }

        [Fact]
        public async Task SaveAs_EnumValue_WritesNumeric()
        {
            var bytes = Xlsx.ToBytes(new[] { new OrderWithEnumDto { Status = Status.Paid } });
            var rows = XlsxIO_TestSupport.ReadSheet(bytes);
            rows[1][1].ShouldBe("1");
        }

        [Fact]
        public async Task SaveAs_DateTimeOffset_AndNullableEnum_WritesNumericAndEmpty()
        {
            var when = new DateTimeOffset(2024, 1, 2, 3, 4, 5, TimeSpan.FromHours(8));
            var bytes = Xlsx.ToBytes(new[]
            {
                new OffsetEnumDto { When = when, OptionalStatus = null }
            });

            var rows = XlsxIO_TestSupport.ReadSheet(bytes);
            rows[1][0].ShouldBe(when.ToString("O", CultureInfo.InvariantCulture));
            rows[1][1].ShouldBeNull();

            var roundTrip = Xlsx.Read<OffsetEnumDto>(new MemoryStream(bytes)).ToList();
            roundTrip.Count.ShouldBe(1);
            roundTrip[0].When.ShouldBe(when);
        }

        [Fact]
        public async Task SaveAsync_NullPath_Throws()
        {
            Should.Throw<ArgumentNullException>(() => Xlsx.Write<OrderDto>((string)null!, Array.Empty<OrderDto>()));
        }

        [Fact]
        public async Task SaveAsync_EmptyPath_Throws()
        {
            Should.Throw<ArgumentException>(() => Xlsx.Write<OrderDto>("", Array.Empty<OrderDto>()));
        }

        [Fact]
        public async Task SaveAs_DisplayAttribute_BecomesHeader()
        {
            var bytes = Xlsx.ToBytes(new[] { new OrderWithFormatDto() });
            var rows = XlsxIO_TestSupport.ReadSheet(bytes);
            rows[0].ShouldContain("Amount");
            rows[0].ShouldContain("Date");
        }

        [Fact]
        public async Task SaveAs_DescriptionAttribute_RoundsTripHeader()
        {
            var bytes = Xlsx.ToBytes(new[] { new DescribedDto { Note = "x" } });
            var rows = XlsxIO_TestSupport.ReadSheet(bytes);
            rows[0].ShouldContain("备注");

            var readBack = Xlsx.Read<DescribedDto>(new MemoryStream(bytes)).ToList();
            readBack.Count.ShouldBe(1);
            readBack[0].Note.ShouldBe("x");
        }

        [Fact]
        public async Task SaveAs_FluentOverridesDisplay()
        {
            var p = new ExportProfile<OrderWithFormatDto>().Column(x => x.Note, c => c.WithName("备注"));
            var bytes = Xlsx.ToBytes(new[] { new OrderWithFormatDto { Note = "x" } }, p);
            var rows = XlsxIO_TestSupport.ReadSheet(bytes);
            rows[0].ShouldContain("备注");
        }

        [Fact]
        public async Task SaveAsBytes_NullableProperties_PreservesNullCells()
        {
            var bytes = Xlsx.ToBytes(new[]
            {
                new NullableDto(),
                new NullableDto
                {
                    Name = "A",
                    Qty = 3,
                    Price = 1.25m,
                    CreatedAt = new DateTime(2024, 1, 2),
                    Enabled = true,
                },
            });

            var rows = XlsxIO_TestSupport.ReadSheet(bytes);
            rows.Count.ShouldBe(3);
            rows[1].All(string.IsNullOrEmpty).ShouldBeTrue();
            rows[2][0].ShouldBe("A");
            rows[2][1].ShouldBe("3");
            rows[2][2].ShouldBe("1.25");
            rows[2][4].ShouldBe("1");
        }

        [Fact]
        public async Task SaveAs_StructType_WritesCorrectly()
        {
            var bytes = Xlsx.ToBytes(new[] { new TestOrderRecord("S1", 5m) });
            var rows = XlsxIO_TestSupport.ReadSheet(bytes);
            rows.Count.ShouldBe(2);
            rows[1][0].ShouldBe("S1");
        }

        [Fact]
        public async Task SaveAs_RecordStruct_WritesCorrectly()
        {
            var bytes = Xlsx.ToBytes(new[] { new TestOrderRecord("S1", 5m) });
            var rows = XlsxIO_TestSupport.ReadSheet(bytes);
            rows.Count.ShouldBe(2);
            rows[1][0].ShouldBe("S1");
        }

        [Fact]
        public async Task SaveAs_WithColumnWidth_WritesColsElement()
        {
            var p = new ExportProfile<OrderDto>().Column(x => x.OrderNo, c => c.WithWidth(25));
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            xml.ShouldContain("<cols>");
            xml.ShouldContain("width=\"25\"");
        }

        [Fact]
        public async Task SaveAs_WithFreezeHeader_WritesSheetView()
        {
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() });
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            xml.ShouldContain("<sheetView");
            xml.ShouldContain("ySplit=\"1\"");
        }

        [Fact]
        public async Task SaveAs_WithoutFreezeHeader_NoSheetView()
        {
            var p = new ExportProfile<OrderDto>().WithFreezeHeader(false);
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            xml.ShouldNotContain("<sheetView");
        }

        [Fact]
        public async Task SaveAsBytes_PackageCriticalXmlParts_AreWellFormed()
        {
            var bytes = Xlsx.ToBytes(new[] { new OrderDto { OrderNo = "S" } });
            var parts = XlsxIO_TestSupport.ReadCriticalXmlParts(bytes);
            foreach (var kv in parts)
            {
                XlsxIO_TestSupport.AssertWellFormedXml(kv.Value, kv.Key);
            }
            parts["xl/worksheets/sheet1.xml"].ShouldContain("S");
            parts["xl/workbook.xml"].ShouldContain("sheets");
        }

        [Fact]
        public async Task SaveMultiSheet_TwoSheets_BothWritten()
        {
            var bytes = Xlsx.WriteWorkbookToBytes(
                new Sheet("First", new[] { new OrderDto { OrderNo = "A" } }),
                new Sheet("Second", new[] { new OrderDto { OrderNo = "B" } }));
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/workbook.xml");
            xml.ShouldContain("First");
            xml.ShouldContain("Second");
        }

        [Fact]
        public async Task WriteMultiSheetAsync_Stream_TwoSheets_BothWritten()
        {
            using var ms = new MemoryStream();
            var bytes = Xlsx.WriteWorkbookToBytes(
                new Sheet("First", new[] { new OrderDto { OrderNo = "A" } }),
                new Sheet("Second", new[] { new OrderDto { OrderNo = "B" } }));
            ms.Write(bytes);
            var xml = XlsxIO_TestSupport.ReadEntry(ms.ToArray(), "xl/workbook.xml");
            xml.ShouldContain("First");
            xml.ShouldContain("Second");
        }

        [Fact]
        public async Task SaveMultiSheetAsync_Path_TwoSheets_BothWritten()
        {
            var path = Path.Combine(Path.GetTempPath(), $"xlsx_multi_test_{Guid.NewGuid():N}.xlsx");
            try
            {
                var bytes = Xlsx.WriteWorkbookToBytes(
                    new Sheet("First", new[] { new OrderDto { OrderNo = "A" } }),
                    new Sheet("Second", new[] { new OrderDto { OrderNo = "B" } }));
                await File.WriteAllBytesAsync(path, bytes);
                var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/workbook.xml");
                xml.ShouldContain("First");
                xml.ShouldContain("Second");
            }
            finally
            {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public async Task WriteMultiSheetAsync_BufferWriter_TwoSheets_BothWritten()
        {
            var output = new ArrayBufferWriter<byte>();
            var bytes = Xlsx.WriteWorkbookToBytes(
                new Sheet("First", new[] { new OrderDto { OrderNo = "A" } }),
                new Sheet("Second", new[] { new OrderDto { OrderNo = "B" } }));
            output.Write(bytes);

            var xml = XlsxIO_TestSupport.ReadEntry(output.WrittenMemory.ToArray(), "xl/workbook.xml");
            xml.ShouldContain("First");
            xml.ShouldContain("Second");
        }

        [Fact]
        public async Task SaveMultiSheet_GenericEmptySheet_PreservesSheetAndHeader()
        {
            var bytes = Xlsx.WriteWorkbookToBytes(
                new Sheet<OrderDto>("EmptyOrders", Array.Empty<OrderDto>()),
                new Sheet<OrderDto>("ActualOrders", new[] { new OrderDto { OrderNo = "A" } }));

            var workbookXml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/workbook.xml");
            workbookXml.ShouldContain("EmptyOrders");
            workbookXml.ShouldContain("ActualOrders");

            var firstSheetXml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            firstSheetXml.ShouldContain("OrderNo");
            firstSheetXml.ShouldContain("Amount");
            firstSheetXml.ShouldContain("CreatedAt");
        }

        [Fact]
        public async Task SaveMultiSheet_NonGenericEmptySheet_PreservesEmptySheetPart()
        {
            var bytes = Xlsx.WriteWorkbookToBytes(
                new Sheet("EmptyOrders", new System.Collections.ArrayList()));

            var workbookXml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/workbook.xml");
            workbookXml.ShouldContain("EmptyOrders");

            var sheetXml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            sheetXml.ShouldContain("<worksheet");
            sheetXml.ShouldNotContain("<sheetData>");
        }

        [Fact]
        public async Task SaveMultiSheet_ZeroSheets_Throws()
        {
            Should.Throw<ArgumentException>(() => { Xlsx.WriteWorkbookToBytes(); });
        }

        [Fact]
        public async Task SaveMultiSheet_NullSheet_Throws()
        {
            Should.Throw<ArgumentNullException>(() => { Xlsx.ToBytes((IEnumerable<OrderDto>)null!); });
        }

        [Fact]
        public async Task SaveMultiSheet_ValidatesSheetNames()
        {
            Should.Throw<ArgumentException>(() => { Xlsx.WriteWorkbookToBytes(new Sheet("Bad/Name", new[] { new OrderDto() })); });
        }

        [Fact]
        public async Task WriteAsync_Stream_WritesValidXlsx()
        {
            using var ms = new MemoryStream();
            Xlsx.Write(ms, new[] { new OrderDto { OrderNo = "S" } });
            ms.Position = 0;
            var items = Xlsx.Read<OrderDto>(ms).ToList();
            items.Count.ShouldBe(1);
            items[0].OrderNo.ShouldBe("S");
        }

        [Fact]
        public async Task WriteAsync_NonSeekableStream_WritesValidXlsx()
        {
            using var inner = new MemoryStream();
            using (var output = new NonSeekableWriteStream(inner))
            {
                Xlsx.Write(output, new[] { new OrderDto { OrderNo = "NS" } });
            }

            inner.Position = 0;
            var items = Xlsx.Read<OrderDto>(inner).ToList();
            items.Count.ShouldBe(1);
            items[0].OrderNo.ShouldBe("NS");
        }

        [Fact]
        public async Task WriteAsync_Stream_WithConfigure_AppliesProfile()
        {
            using var ms = new MemoryStream();
            Xlsx.Write(ms, new[] { new OrderDto() }, p => p
                .Column(x => x.OrderNo, c => c.WithName("订单")));

            var xml = XlsxIO_TestSupport.ReadEntry(ms.ToArray(), "xl/worksheets/sheet1.xml");
            xml.ShouldContain("订单");
        }

        [Fact]
        public async Task WriteAsync_BufferWriter_AsyncEnumerable_WritesValidXlsx()
        {
            async IAsyncEnumerable<OrderDto> Data()
            {
                yield return new OrderDto { OrderNo = "AsyncBufferWriter" };
                await Task.Yield();
            }

            using var ms = new MemoryStream();
            await Xlsx.WriteAsync(ms, Data());
            var rows = XlsxIO_TestSupport.ReadSheet(ms.ToArray());
            rows[1][0].ShouldBe("AsyncBufferWriter");
        }

        [Fact]
        public async Task WriteAsync_AsyncEnumerable_WithOptions_WritesValidXlsx()
        {
            async IAsyncEnumerable<OrderDto> Data()
            {
                yield return new OrderDto { OrderNo = "AsyncWrite" };
                await Task.Yield();
            }

            using var ms = new MemoryStream();
            await Xlsx.WriteAsync(ms, Data(), p => p.Sheet("AsyncOrders"));

            var rows = XlsxIO_TestSupport.ReadSheet(ms.ToArray());
            rows[1][0].ShouldBe("AsyncWrite");
        }

        [Fact]
        public async Task WriteAsync_AsyncEnumerable_AppliesRowFilter()
        {
            async IAsyncEnumerable<OrderDto> Data()
            {
                yield return new OrderDto { OrderNo = "keep-1" };
                await Task.Yield();
                yield return new OrderDto { OrderNo = "drop" };
                yield return new OrderDto { OrderNo = "keep-2" };
            }

            using var ms = new MemoryStream();
            await Xlsx.WriteAsync(ms, Data(), new ExportProfile<OrderDto>()
                .Where(x => x.OrderNo != "drop"));

            var rows = XlsxIO_TestSupport.ReadSheet(ms.ToArray());
            rows.Count.ShouldBe(3);
            rows[1][0].ShouldBe("keep-1");
            rows[2][0].ShouldBe("keep-2");
        }

        [Fact]
        public async Task WriteAsync_Iterator_AppliesRowFilter()
        {
            IEnumerable<OrderDto> Data()
            {
                yield return new OrderDto { OrderNo = "keep-1" };
                yield return new OrderDto { OrderNo = "drop" };
                yield return new OrderDto { OrderNo = "keep-2" };
            }

            using var ms = new MemoryStream();
            await Xlsx.WriteAsync(ms, Data(), new ExportProfile<OrderDto>()
                .Where(x => x.OrderNo != "drop"));

            var rows = XlsxIO_TestSupport.ReadSheet(ms.ToArray());
            rows.Count.ShouldBe(3);
            rows[1][0].ShouldBe("keep-1");
            rows[2][0].ShouldBe("keep-2");
        }


        [Fact]
        public async Task WriteMultiSheetAsync_Stream_WritesValidWorkbook()
        {
            using var ms = new MemoryStream();
            var bytes = Xlsx.WriteWorkbookToBytes(
                new Sheet<OrderDto>("S1", new[] { new OrderDto { OrderNo = "A" } }),
                new Sheet<OrderDto>("S2", Array.Empty<OrderDto>()));
            ms.Write(bytes);
            var workbookXml = XlsxIO_TestSupport.ReadEntry(ms.ToArray(), "xl/workbook.xml");
            workbookXml.ShouldContain("S1");
            workbookXml.ShouldContain("S2");
        }

        [Fact]
        public async Task WriteMultiSheetAsync_BufferWriter_WithCompression_WritesValidWorkbook()
        {
            var output = new ArrayBufferWriter<byte>();
            var bytes = Xlsx.WriteWorkbookToBytes(
                new Sheet("First", new[] { new OrderDto { OrderNo = "A" } }),
                new Sheet("Second", new[] { new OrderDto { OrderNo = "B" } }));
            output.Write(bytes);

            var xml = XlsxIO_TestSupport.ReadEntry(output.WrittenMemory.ToArray(), "xl/workbook.xml");
            xml.ShouldContain("First");
            xml.ShouldContain("Second");
        }

        [Fact]
        public async Task SaveAsBytesAsync_WithSheetName_DoesNotCreateDuplicateSheets()
        {
            async IAsyncEnumerable<OrderDto> Data()
            {
                yield return new OrderDto { OrderNo = "A" };
                await Task.CompletedTask;
            }

            using var ms = new MemoryStream();
            await Xlsx.WriteAsync(ms, Data(), p => p.Sheet("Orders"));
            var bytes = ms.ToArray();
            var workbookXml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/workbook.xml");
            var sheetTagCount = (workbookXml.Length - workbookXml.Replace("<sheet ", string.Empty, StringComparison.Ordinal).Length) / "<sheet ".Length;

            sheetTagCount.ShouldBe(1);
            workbookXml.ShouldContain("Orders");
        }

        [Fact]
        public async Task SaveAsBytes_NullData_ThrowsArgumentNull()
        {
            Should.Throw<ArgumentNullException>(() => { Xlsx.ToBytes<OrderDto>((IEnumerable<OrderDto>)null!); });
        }

        [Fact]
        public async Task SaveAsBytes_EmptyEnumerable_WritesValidXlsx()
        {
            var bytes = Xlsx.ToBytes(Array.Empty<OrderDto>());
            var rows = XlsxIO_TestSupport.ReadSheet(bytes);
            rows.Count.ShouldBe(1);
        }

        private sealed class NonSeekableWriteStream : Stream
        {
            private readonly Stream _inner;

            public NonSeekableWriteStream(Stream inner)
            {
                _inner = inner;
            }

            public override bool CanRead => false;
            public override bool CanSeek => false;
            public override bool CanWrite => true;
            public override long Length => throw new NotSupportedException();

            public override long Position
            {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }

            public override void Flush() => _inner.Flush();

            public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

            public override void SetLength(long value) => _inner.SetLength(value);

            public override void Write(byte[] buffer, int offset, int count) => _inner.Write(buffer, offset, count);

#if NETCOREAPP3_1_OR_GREATER
            public override void Write(ReadOnlySpan<byte> buffer) => _inner.Write(buffer);
#endif
        }
    }
}
