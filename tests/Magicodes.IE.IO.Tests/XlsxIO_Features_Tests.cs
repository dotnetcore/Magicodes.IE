
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
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
        public async Task MergeCells_WritesMergeCellsElement()
        {
            var p = new ExportProfile<OrderDto>().WithMergeCells("A1:B2");
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            xml.ShouldContain("<mergeCells");
        }

        [Fact]
        public async Task AutoFilter_WritesAutoFilterElement()
        {
            var p = new ExportProfile<OrderDto>().WithAutoFilter("A1:C10");
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            xml.ShouldContain("<autoFilter ref=\"A1:C10\"/>");
        }

        [Fact]
        public async Task Hyperlink_WritesHyperlinksElement()
        {
            var p = new ExportProfile<OrderDto>().WithHyperlink("A2", "https://example.com");
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            xml.ShouldContain("<hyperlinks>");
            xml.ShouldContain("r:id=\"rIdH1\"");
            var rels = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/_rels/sheet1.xml.rels");
            rels.ShouldContain("https://example.com");
        }

        [Fact]
        public async Task Hyperlink_MultipleLinks_WritesAll()
        {
            var p = new ExportProfile<OrderDto>()
                .WithHyperlink("A2", "https://a.com")
                .WithHyperlink("B2", "https://b.com");
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var rels = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/_rels/sheet1.xml.rels");
            rels.ShouldContain("https://a.com");
            rels.ShouldContain("https://b.com");
        }

        private static readonly byte[] TinyPng = new byte[]
        {
            0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
            0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41,
            0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
            0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
            0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
            0x42, 0x60, 0x82
        };

        [Fact]
        public async Task AddImage_WritesImageAndDrawing()
        {
            using var ms = new MemoryStream();
            using (var w = new XlsxWriter(ms))
            {
                w.AddSheet("S1");
                w.AddImage(TinyPng, "png", "A1", "B2");
            }
            using var zip = new ZipArchive(new MemoryStream(ms.ToArray()), ZipArchiveMode.Read);
            zip.GetEntry("xl/drawings/drawing1.xml").ShouldNotBeNull();
            zip.GetEntry("xl/media/image1.png").ShouldNotBeNull();
        }

        [Fact]
        public async Task AddImage_ExcelParseable_RelsAndNamespaces()
        {
            using var ms = new MemoryStream();
            using (var w = new XlsxWriter(ms))
            {
                w.AddSheet("S1");
                w.AddImage(TinyPng, "png", "A1", "B2");
            }
            using var zip = new ZipArchive(new MemoryStream(ms.ToArray()), ZipArchiveMode.Read);

            var drawRelsEntry = zip.GetEntry("xl/drawings/_rels/drawing1.xml.rels");
            drawRelsEntry.ShouldNotBeNull();
            using (var s = drawRelsEntry!.Open())
            using (var sr = new StreamReader(s))
            {
                var drawRels = sr.ReadToEnd();
                drawRels.ShouldContain("Id=\"rId1\"");
                drawRels.ShouldContain("Target=\"../media/image1.png\"");
            }

            var drawEntry = zip.GetEntry("xl/drawings/drawing1.xml");
            drawEntry.ShouldNotBeNull();
            using (var s = drawEntry!.Open())
            using (var sr = new StreamReader(s))
            {
                var draw = sr.ReadToEnd();
                draw.ShouldContain("xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"");
                draw.ShouldContain("r:embed=\"rId1\"");
            }

            var sheetRelsEntry = zip.GetEntry("xl/worksheets/_rels/sheet1.xml.rels");
            sheetRelsEntry.ShouldNotBeNull();
            using (var s = sheetRelsEntry!.Open())
            using (var sr = new StreamReader(s))
            {
                var sheetRels = sr.ReadToEnd();
                sheetRels.ShouldContain("Id=\"rIdImage1\"");
                sheetRels.ShouldContain("Target=\"../drawings/drawing1.xml\"");
            }

            var sheetEntry = zip.GetEntry("xl/worksheets/sheet1.xml");
            sheetEntry.ShouldNotBeNull();
            using (var s = sheetEntry!.Open())
            using (var sr = new StreamReader(s))
            {
                var sheet = sr.ReadToEnd();
                sheet.ShouldContain("<drawing r:id=\"rIdImage1\"/>");
            }
        }

        [Fact]
        public void AddImage_InvalidExtension_Throws()
        {
            using var ms = new MemoryStream();
            using var writer = new XlsxWriter(ms);
            writer.AddSheet("S1");
            Should.Throw<ArgumentException>(() => writer.AddImage(new byte[] { 1, 2, 3 }, "tiff", "A1", "B2"));
        }

        [Fact]
        public void AddImage_InvalidAnchor_Throws()
        {
            using var ms = new MemoryStream();
            using var writer = new XlsxWriter(ms);
            writer.AddSheet("S1");
            Should.Throw<ArgumentException>(() => writer.AddImage(new byte[] { 1, 2, 3 }, "png", "INVALID", "B2"));
        }

        [Fact]
        public void AddHyperlink_JavascriptScheme_Throws()
        {
            using var ms = new MemoryStream();
            using var writer = new XlsxWriter(ms);
            writer.AddSheet("S1");
            Should.Throw<ArgumentException>(() => writer.AddHyperlink("A1", "javascript:alert(1)"));
        }

        [Fact]
        public void AddHyperlink_HttpUrl_Accepted()
        {
            using var ms = new MemoryStream();
            using var writer = new XlsxWriter(ms);
            writer.AddSheet("S1");
            Should.NotThrow(() => writer.AddHyperlink("A1", "https://example.com"));
        }

        [Fact]
        public void AddComment_MultipleAuthors_SerializesStableAuthorIds()
        {
            using var ms = new MemoryStream();
            using (var writer = new XlsxWriter(ms))
            {
                writer.AddSheet("S1");
                writer.AddComment(new Comment(0, 0, "Alice", "first by alice"));
                writer.AddComment(new Comment(1, 0, "Bob",   "first by bob"));
                writer.AddComment(new Comment(2, 0, "Alice", "second by alice"));
                writer.AddComment(new Comment(3, 0, "Bob",   "second by bob"));
            }

            using var za = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: false);
            var entry = za.GetEntry("xl/comments1.xml");
            entry.ShouldNotBeNull();
            using var es = entry!.Open();
            var doc = new XmlDocument();
            using (var xr = XmlReader.Create(es, new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null }))
                doc.Load(xr);

            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            var authors = doc.SelectNodes("/x:comments/x:authors/x:author", ns);
            authors!.Count.ShouldBe(2);
            authors[0]!.InnerText.ShouldBe("Alice");
            authors[1]!.InnerText.ShouldBe("Bob");

            var commentList = doc.SelectNodes("/x:comments/x:commentList/x:comment", ns);
            commentList!.Count.ShouldBe(4);
            commentList[0]!.Attributes!["ref"]!.Value.ShouldBe("A1");
            commentList[1]!.Attributes!["ref"]!.Value.ShouldBe("A2");
            commentList[2]!.Attributes!["ref"]!.Value.ShouldBe("A3");
            commentList[3]!.Attributes!["ref"]!.Value.ShouldBe("A4");
            commentList[0]!.Attributes!["authorId"]!.Value.ShouldBe("0");
            commentList[1]!.Attributes!["authorId"]!.Value.ShouldBe("1");
            commentList[2]!.Attributes!["authorId"]!.Value.ShouldBe("0");
            commentList[3]!.Attributes!["authorId"]!.Value.ShouldBe("1");
        }

        [Fact]
        public void AddComment_SingleAuthor_AllCommentsAuthorIdZero()
        {
            using var ms = new MemoryStream();
            using (var writer = new XlsxWriter(ms))
            {
                writer.AddSheet("S1");
                writer.AddComment(new Comment(0, 0, "Solo", "a"));
                writer.AddComment(new Comment(0, 1, "Solo", "b"));
                writer.AddComment(new Comment(0, 2, "Solo", "c"));
            }
            using var za = new ZipArchive(ms, ZipArchiveMode.Read);
            using var es = za.GetEntry("xl/comments1.xml")!.Open();
            var doc = new XmlDocument();
            doc.Load(es);
            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var comments = doc.SelectNodes("/x:comments/x:commentList/x:comment", ns);
            for (int i = 0; i < comments!.Count; i++)
                comments[i]!.Attributes!["authorId"]!.Value.ShouldBe("0");

            comments[0]!.Attributes!["ref"]!.Value.ShouldBe("A1");
            comments[1]!.Attributes!["ref"]!.Value.ShouldBe("B1");
            comments[2]!.Attributes!["ref"]!.Value.ShouldBe("C1");
        }

        [Fact]
        public void AddComment_WritesWorksheetRelationship()
        {
            using var ms = new MemoryStream();
            using (var writer = new XlsxWriter(ms))
            {
                writer.AddSheet("S1");
                writer.AddComment(new Comment(0, 0, "Solo", "comment"));
            }

            using var za = new ZipArchive(ms, ZipArchiveMode.Read, leaveOpen: false);
            var relsEntry = za.GetEntry("xl/worksheets/_rels/sheet1.xml.rels");
            relsEntry.ShouldNotBeNull();
            using var es = relsEntry!.Open();
            var rels = new StreamReader(es).ReadToEnd();
            rels.ShouldContain("Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\"");
            rels.ShouldContain("Target=\"../comments1.xml\"");
            rels.ShouldContain("Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\"");
            rels.ShouldContain("Target=\"../drawings/vmlDrawing1.vml\"");

            var sheetXml = XlsxIO_TestSupport.ReadEntry(ms.ToArray(), "xl/worksheets/sheet1.xml");
            sheetXml.ShouldContain("<legacyDrawing r:id=\"rIdVml1\"/>");

            var vmlEntry = za.GetEntry("xl/drawings/vmlDrawing1.vml");
            vmlEntry.ShouldNotBeNull();
            using (var vmlStream = vmlEntry!.Open())
            using (var reader = new StreamReader(vmlStream))
            {
                var vml = reader.ReadToEnd();
                vml.ShouldContain("<x:Row>0</x:Row>");
                vml.ShouldContain("<x:Column>0</x:Column>");
            }
        }

        [Fact]
        public async Task DataValidation_List_WritesDataValidationElement()
        {
            using var ms = new MemoryStream();
            using (var writer = new XlsxWriter(ms))
            {
                writer.AddSheet("S1");
                writer.AddDataValidation(new DataValidation("A2:A10", DataValidationType.List, "\"选项1,选项2\""));
                writer.WriteHeader(XlsxIO_TestSupport.MakeCols());
                writer.WriteRows(new[] { new XlsxIO_TestSupport.Order { A = 1 } }, XlsxIO_TestSupport.MakeTypedPlan(new[] { "A" }));
            }
            var xml = XlsxIO_TestSupport.ReadEntry(ms.ToArray(), "xl/worksheets/sheet1.xml");
            xml.ShouldContain("<dataValidation");
            xml.ShouldContain("type=\"list\"");
        }

        [Fact]
        public async Task DataValidation_IntegerRange_WritesCorrectType()
        {
            using var ms = new MemoryStream();
            using (var writer = new XlsxWriter(ms))
            {
                writer.AddSheet("S1");
                writer.AddDataValidation(new DataValidation("B2:B100", DataValidationType.Integer, "1", "100",
                    op: DataValidationOperator.Between));
                writer.WriteHeader(XlsxIO_TestSupport.MakeCols());
                writer.WriteRows(new[] { new XlsxIO_TestSupport.Order { A = 1 } }, XlsxIO_TestSupport.MakeTypedPlan(new[] { "A" }));
            }
            var xml = XlsxIO_TestSupport.ReadEntry(ms.ToArray(), "xl/worksheets/sheet1.xml");
            xml.ShouldContain("type=\"whole\"");
            xml.ShouldContain("operator=\"between\"");
        }

        [Fact]
        public async Task DataValidation_List_WithExportProfile()
        {
            var p = new ExportProfile<OrderDto>().Column(x => x.OrderNo, c => c.WithName("Status"));
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            xml.ShouldContain("Status");
        }

        [Fact]
        public async Task Formula_WritesFormulaElement()
        {
            var p = new ExportProfile<OrderWithFormatDto>()
                .Column(x => x.Amount, c => c.WithFormula("SUM(A1:A10)"));
            var bytes = Xlsx.ToBytes(new[] { new OrderWithFormatDto { Amount = 5 } }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            xml.ShouldContain("<f>");
            xml.ShouldContain("SUM(A1:A10)");
        }

        [Fact]
        public async Task Formula_ColumnWithRowPlaceholder_WritesRowAwareFormula()
        {
            var p = new ExportProfile<OrderWithFormatDto>()
                .Column(x => x.Amount, c => c.WithFormula("A{row}*2"));
            var bytes = Xlsx.ToBytes(
                new[] { new OrderWithFormatDto { Amount = 1 }, new OrderWithFormatDto { Amount = 2 } }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            xml.ShouldContain("A2*2");
            xml.ShouldContain("A3*2");
        }

        [Fact]
        public async Task Formula_MultipleRowPlaceholders_AllReplaced()
        {
            var p = new ExportProfile<OrderWithFormatDto>()
                .Column(x => x.Amount, c => c.WithFormula("A{row}+B{row}"));
            var bytes = Xlsx.ToBytes(new[] { new OrderWithFormatDto { Amount = 1 } }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            xml.ShouldContain("A2+B2");
        }

        [Fact]
        public async Task SetDefaultRowHeight_AppearsInSheetFormatPr()
        {
            using var ms = new MemoryStream();
            using (var writer = new XlsxWriter(ms, defaultRowHeight: 22.5))
            {
                writer.AddSheet("S1");
                var cols = XlsxIO_TestSupport.MakeCols();
                writer.WriteSheetMeta(cols, freezeHeader: true);
                writer.WriteHeader(cols);
                writer.WriteRows(new[] { new XlsxIO_TestSupport.Order { A = 1 } }, XlsxIO_TestSupport.MakeTypedPlan(new[] { "A" }));
            }
            var xml = XlsxIO_TestSupport.ReadEntry(ms.ToArray(), "xl/worksheets/sheet1.xml");
            xml.ShouldContain("defaultRowHeight=\"16.875\"");
        }

        [Fact]
        public async Task SetNextRowHeight_EmitsHtAttrOnNextRow()
        {
            using var ms = new MemoryStream();
            using (var writer = new XlsxWriter(ms))
            {
                writer.AddSheet("S1");
                writer.WriteHeader(XlsxIO_TestSupport.MakeCols());
                writer.SetNextRowHeight(33.3);
                writer.WriteRows(new[] { new XlsxIO_TestSupport.Order { A = 1 } }, XlsxIO_TestSupport.MakeTypedPlan(new[] { "A" }));
            }
            var xml = XlsxIO_TestSupport.ReadEntry(ms.ToArray(), "xl/worksheets/sheet1.xml");
            // Double formatting differs across runtimes (net471 G15 vs net6+ shortest round-trip),
            // so compare the parsed numeric value with a tolerance instead of an exact string.
            var htMatch = System.Text.RegularExpressions.Regex.Match(xml, "ht=\"([^\"]*)\"");
            htMatch.Success.ShouldBeTrue();
            double.TryParse(htMatch.Groups[1].Value, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out var ht).ShouldBeTrue();
            ht.ShouldBe(24.975, 1e-6);
            xml.ShouldContain("customHeight=\"1\"");
        }

        [Fact]
        public async Task EnableSharedStrings_RepeatedStrings_AreDeduplicated()
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
                writer.WriteRows(Enumerable.Range(0, 100).Select(_ => new StringDto { Name = "重复" }), plan);
            }
            using var za = new ZipArchive(ms, ZipArchiveMode.Read);
            var sstEntry = za.GetEntry("xl/sharedStrings.xml");
            sstEntry.ShouldNotBeNull();
            string sstXml = XlsxIO_TestSupport.ReadEntry(ms.ToArray(), "xl/sharedStrings.xml");
            int sstCount = (sstXml.Length - sstXml.Replace("<si>", string.Empty).Length) / 4;
            sstCount.ShouldBe(1);
        }

        [Fact]
        public async Task EnableSharedStrings_SstCountEqualsTotalCellRefs()
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
                writer.WriteRows(Enumerable.Range(0, 100).Select(i => new StringDto { Name = i % 2 == 0 ? "a" : "b" }), plan);
            }
            var sstXml = XlsxIO_TestSupport.ReadEntry(ms.ToArray(), "xl/sharedStrings.xml");
            sstXml.ShouldContain("count=\"100\"");
            sstXml.ShouldContain("uniqueCount=\"2\"");
        }

        [Fact]
        public async Task SetOutline_AppearsInSheetFormatPr()
        {
            using var ms = new MemoryStream();
            using (var writer = new XlsxWriter(ms))
            {
                writer.AddSheet("S1");
                writer.SetOutline(new OutlineSettings { SummaryBelow = false, SummaryRight = false });
                var cols = XlsxIO_TestSupport.MakeCols();
                writer.WriteSheetMeta(cols, freezeHeader: true);
                writer.WriteHeader(cols);
                writer.WriteRows(new[] { new XlsxIO_TestSupport.Order { A = 1 } }, XlsxIO_TestSupport.MakeTypedPlan(new[] { "A" }));
            }
            var xml = XlsxIO_TestSupport.ReadEntry(ms.ToArray(), "xl/worksheets/sheet1.xml");
            xml.ShouldContain("summaryBelow=\"0\"");
            xml.ShouldContain("summaryRight=\"0\"");
            System.Globalization.CultureInfo.InvariantCulture.CompareInfo.IndexOf(xml, "<sheetPr", System.Globalization.CompareOptions.Ordinal)
                .ShouldBeLessThan(System.Globalization.CultureInfo.InvariantCulture.CompareInfo.IndexOf(xml, "<sheetViews", System.Globalization.CompareOptions.Ordinal));
        }

        [Fact]
        public async Task AddNamedRange_WorkbookLevel_AppearsInWorkbookXml()
        {
            using var ms = new MemoryStream();
            using (var writer = new XlsxWriter(ms))
            {
                writer.AddSheet("S1");
                writer.AddNamedRange(new NamedRange("MyRange", "S1!$A$1:$A$10"));
                writer.WriteHeader(XlsxIO_TestSupport.MakeCols());
                writer.WriteRows(new[] { new XlsxIO_TestSupport.Order { A = 1 } }, XlsxIO_TestSupport.MakeTypedPlan(new[] { "A" }));
            }
            var xml = XlsxIO_TestSupport.ReadEntry(ms.ToArray(), "xl/workbook.xml");
            xml.ShouldContain("<definedNames>");
            xml.ShouldContain("name=\"MyRange\"");
            xml.ShouldContain("S1!$A$1:$A$10");
        }

        [Fact]
        public async Task AddConditionalFormatting_AppearsInSheetXml()
        {
            using var ms = new MemoryStream();
            using (var writer = new XlsxWriter(ms))
            {
                writer.AddSheet("S1");
                writer.AddConditionalFormatting(new ConditionalFormatting("A1:A100",
                    new CfRule(CfOperator.GreaterThan, "100")));
                writer.WriteHeader(XlsxIO_TestSupport.MakeCols());
                writer.WriteRows(new[] { new XlsxIO_TestSupport.Order { A = 1 } }, XlsxIO_TestSupport.MakeTypedPlan(new[] { "A" }));
            }
            var xml = XlsxIO_TestSupport.ReadEntry(ms.ToArray(), "xl/worksheets/sheet1.xml");
            xml.ShouldContain("<conditionalFormatting");
            xml.ShouldContain("sqref=\"A1:A100\"");
            xml.ShouldContain("<cfRule");
            xml.ShouldContain("type=\"cellIs\"");
            xml.ShouldContain("operator=\"greaterThan\"");
        }

        [Fact]
        public async Task MergeCells_AppearsInSheetXml()
        {
            using var ms = new MemoryStream();
            using (var writer = new XlsxWriter(ms))
            {
                writer.AddSheet("S1");
                writer.MergeCells("A1:B2");
                writer.WriteHeader(XlsxIO_TestSupport.MakeCols());
                writer.WriteRows(new[] { new XlsxIO_TestSupport.Order { A = 1 } }, XlsxIO_TestSupport.MakeTypedPlan(new[] { "A" }));
            }
            var xml = XlsxIO_TestSupport.ReadEntry(ms.ToArray(), "xl/worksheets/sheet1.xml");
            xml.ShouldContain("<mergeCells count=\"1\">");
            xml.ShouldContain("ref=\"A1:B2\"");
        }

        [Fact]
        public async Task SetAutoFilter_AppearsInSheetXml()
        {
            using var ms = new MemoryStream();
            using (var writer = new XlsxWriter(ms))
            {
                writer.AddSheet("S1");
                writer.WriteHeader(XlsxIO_TestSupport.MakeCols());
                writer.WriteRows(new[] { new XlsxIO_TestSupport.Order { A = 1 } }, XlsxIO_TestSupport.MakeTypedPlan(new[] { "A" }));
                writer.SetAutoFilter("A1:A10");
            }
            var xml = XlsxIO_TestSupport.ReadEntry(ms.ToArray(), "xl/worksheets/sheet1.xml");
            xml.ShouldContain("<autoFilter ref=\"A1:A10\"/>");
        }

        [Fact]
        public async Task AutoSst_RepeatedStrings_EnablesSharedStrings()
        {
            var data = Enumerable.Range(0, 100).Select(i => new StringDto { Name = "A" }).ToList();
            var p = new ExportProfile<StringDto>().WithAutoSst();
            var bytes = Xlsx.ToBytes(data, p);
            using var za = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
            za.GetEntry("xl/sharedStrings.xml").ShouldNotBeNull();
        }

        [Fact]
        public async Task AutoSst_UniqueStrings_DoesNotEnableSharedStrings()
        {
            var data = Enumerable.Range(0, 100).Select(i => new StringDto { Name = $"unique-{i}" }).ToList();
            var p = new ExportProfile<StringDto>().WithAutoSst();
            var bytes = Xlsx.ToBytes(data, p);
            using var za = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
            za.GetEntry("xl/sharedStrings.xml").ShouldBeNull();
        }

        [Fact]
        public void AutoSst_RowFilterStillDetectsSharedStrings()
        {
            var data = Enumerable.Range(0, 100).Select(_ => new StringDto { Name = "A" });
            var p = new ExportProfile<StringDto>().Where(_ => true).WithAutoSst();
            var bytes = Xlsx.ToBytes(data, p);

            using var za = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
            za.GetEntry("xl/sharedStrings.xml").ShouldNotBeNull();
        }

        [Fact]
        public async Task AutoSst_AsyncEnumerableUsesDetectedMode()
        {
            async IAsyncEnumerable<StringDto> Data()
            {
                for (int i = 0; i < 100; i++)
                    yield return new StringDto { Name = "A" };
            }

            using var ms = new MemoryStream();
            await Xlsx.WriteAsync(ms, Data(), new ExportProfile<StringDto>().WithAutoSst());

            using var za = new ZipArchive(new MemoryStream(ms.ToArray()), ZipArchiveMode.Read);
            za.GetEntry("xl/sharedStrings.xml").ShouldNotBeNull();
        }

        [Fact]
        public async Task AutoSst_NoCompression_StaysInlineString()
        {
            var data = Enumerable.Range(0, 100).Select(i => new StringDto { Name = "A" }).ToList();
            var p = new ExportProfile<StringDto>().WithAutoSst();
            var bytes = Xlsx.ToBytes(data, p, new XlsxWriteOptions { Compression = CompressionLevel.NoCompression });
            using var za = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
            za.GetEntry("xl/sharedStrings.xml").ShouldBeNull();
        }

        [Fact]
        public async Task AutoSst_SharedStrings_PreservesWhitespaceAndEscapesXml()
        {
            var data = Enumerable.Range(0, 100).Select(_ => new StringDto { Name = " <a>&\"' " }).ToList();
            var bytes = Xlsx.ToBytes(data, new ExportProfile<StringDto>().WithAutoSst());
            var sstXml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/sharedStrings.xml");

            sstXml.ShouldContain("xml:space=\"preserve\"");
            sstXml.ShouldContain("&lt;a&gt;&amp;&quot;&apos;");
        }

        [Fact]
        public async Task WriteInlineStringCell_AsciiOnly_ProducesSpecCompliantBytes()
        {
            var ascii = Enumerable.Range(0, 100).Select(i => new StringDto { Name = $"name-{i}" }).ToList();
            var bytes = Xlsx.ToBytes(ascii, new ExportProfile<StringDto>());
            bytes.ShouldNotBeEmpty();
            string sheetXml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            sheetXml.ShouldContain("r=\"A1\" t=\"inlineStr\"><is><t>Name");
            sheetXml.ShouldContain("r=\"A2\" t=\"inlineStr\"><is><t>name-0</t></is>");
            sheetXml.ShouldContain("r=\"A101\" t=\"inlineStr\"><is><t>name-99</t></is>");
            sheetXml.ShouldNotContain("name-0&lt;");
            sheetXml.ShouldNotContain("name-0&amp;");
        }

        [Fact]
        public async Task WriteInlineStringCell_AsciiWhitespace_PreservesXmlSpace()
        {
            var bytes = Xlsx.ToBytes(new[] { new StringDto { Name = " leading trailing " } }, new ExportProfile<StringDto>());
            string sheetXml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");

            sheetXml.ShouldContain("<is><t xml:space=\"preserve\"> leading trailing </t></is>");
        }

        [Fact]
        public async Task WriteInlineStringCell_NonAscii_ProducesSpecCompliantBytes()
        {
            var data = Enumerable.Range(0, 100).Select(i => new StringDto { Name = $"中文-{i}" }).ToList();
            var bytes = Xlsx.ToBytes(data, new ExportProfile<StringDto>());
            string sheetXml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            sheetXml.ShouldContain("<is><t>中文-0</t></is>");
            sheetXml.ShouldContain("<is><t>中文-99</t></is>");
        }

        [Fact]
        public async Task WriteInlineStringCell_XmlSpecialChars_AreEscaped()
        {
            var data = new[] {
                new StringDto { Name = "a&b<c>d\"e'f" },
                new StringDto { Name = "plain-ascii-with-no-special" },
            };
            var bytes = Xlsx.ToBytes(data, new ExportProfile<StringDto>());
            string sheetXml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            sheetXml.ShouldContain("a&amp;b&lt;c&gt;d&quot;e&apos;f");
            sheetXml.ShouldContain("<is><t>plain-ascii-with-no-special</t></is>");
            sheetXml.ShouldNotContain("plain-ascii-with-no-special&amp;");
            sheetXml.ShouldNotContain("plain-ascii-with-no-special&lt;");
        }

        [Fact]
        public async Task WriteInlineStringCell_LargeAscii_TakesFastPath()
        {
            var data = Enumerable.Range(0, 100).Select(i => new StringDto
            {
                Name = new string('a', 30) + i.ToString("D3")
            }).ToList();
            var bytes = Xlsx.ToBytes(data, new ExportProfile<StringDto>());
            string sheetXml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            sheetXml.ShouldContain("r=\"A2\" t=\"inlineStr\"><is><t>aaaa");
            sheetXml.ShouldContain("r=\"A101\" t=\"inlineStr\"><is><t>aaa");
            sheetXml.ShouldContain("099</t></is></c></row>");
        }


        [Fact]
        public async Task SingleSheet_SheetProtection_WrittenToFinalXml()
        {
            using var ms = new MemoryStream();
            using (var writer = new XlsxWriter(ms))
            {
                writer.AddSheet("Sheet1");
                writer.SetSheetProtection(new SheetProtection
                {
                    FormatCells = false,
                    Sort = false,
                    AutoFilter = false,
                    PasswordHash = "ABC123",
                });
                writer.WriteHeader(XlsxIO_TestSupport.MakeCols());
                writer.WriteRows(new[] { new XlsxIO_TestSupport.Order { A = 1 } }, XlsxIO_TestSupport.MakeTypedPlan(new[] { "A" }));
            }
            var xml = XlsxIO_TestSupport.ReadEntry(ms.ToArray(), "xl/worksheets/sheet1.xml");
            xml.ShouldContain("<sheetProtection");
            xml.ShouldContain("formatCells=\"0\"");
            xml.ShouldContain("sort=\"0\"");
            xml.ShouldContain("autoFilter=\"0\"");
            xml.ShouldContain("password=\"ABC123\"");
        }

        [Fact]
        public void SheetProtection_PrecedesAutoFilterInWorksheetXml()
        {
            using var ms = new MemoryStream();
            using (var writer = new XlsxWriter(ms))
            {
                writer.AddSheet("Sheet1");
                writer.SetSheetProtection(new SheetProtection());
                writer.SetAutoFilter("A1:A2");
                writer.WriteHeader(XlsxIO_TestSupport.MakeCols());
            }

            var xml = XlsxIO_TestSupport.ReadEntry(ms.ToArray(), "xl/worksheets/sheet1.xml");
            System.Globalization.CultureInfo.InvariantCulture.CompareInfo.IndexOf(xml, "<sheetProtection", System.Globalization.CompareOptions.Ordinal)
                .ShouldBeLessThan(
                    System.Globalization.CultureInfo.InvariantCulture.CompareInfo.IndexOf(xml, "<autoFilter", System.Globalization.CompareOptions.Ordinal));
        }

        [Fact]
        public async Task SingleSheet_PageSetup_WrittenToFinalXml()
        {
            using var ms = new MemoryStream();
            using (var writer = new XlsxWriter(ms))
            {
                writer.AddSheet("Sheet1");
                writer.SetPageSetup(new PageSetup
                {
                    Orientation = "landscape",
                    OddHeader = "&LHeader <1>",
                    OddFooter = "&RPage &P of &N",
                    EvenHeader = "Even header",
                    EvenFooter = "Even footer",
                });
                writer.WriteHeader(XlsxIO_TestSupport.MakeCols());
                writer.WriteRows(new[] { new XlsxIO_TestSupport.Order { A = 1 } }, XlsxIO_TestSupport.MakeTypedPlan(new[] { "A" }));
            }
            var xml = XlsxIO_TestSupport.ReadEntry(ms.ToArray(), "xl/worksheets/sheet1.xml");
            xml.ShouldContain("<pageSetup");
            xml.ShouldContain("<headerFooter differentOddEven=\"1\">");
            xml.ShouldContain("&amp;LHeader &lt;1&gt;");
            xml.ShouldContain("<oddFooter>&amp;RPage &amp;P of &amp;N</oddFooter>");
        }

        [Fact]
        public async Task SingleSheet_FitToPage_EmitsPageSetUpPr()
        {
            // fitToWidth/fitToHeight on <pageSetup> are ignored by Excel unless
            // <sheetPr><pageSetUpPr fitToPage="1"/> is also emitted.
            using var ms = new MemoryStream();
            using (var writer = new XlsxWriter(ms))
            {
                writer.AddSheet("Sheet1");
                writer.SetPageSetup(new PageSetup { FitToWidth = 1, FitToHeight = 0 });
                var cols = XlsxIO_TestSupport.MakeCols();
                writer.WriteSheetMeta(cols, freezeHeader: false);
                writer.WriteHeader(cols);
                writer.WriteRows(new[] { new XlsxIO_TestSupport.Order { A = 1 } }, XlsxIO_TestSupport.MakeTypedPlan(new[] { "A" }));
            }
            var xml = XlsxIO_TestSupport.ReadEntry(ms.ToArray(), "xl/worksheets/sheet1.xml");
            xml.ShouldContain("<sheetPr>");
            xml.ShouldContain("<pageSetUpPr fitToPage=\"1\"/>");
            xml.ShouldContain("fitToWidth=\"1\"");
        }

        [Fact]
        public async Task SingleSheet_Table_WrittenToFinalXml()
        {
            using var ms = new MemoryStream();
            using (var writer = new XlsxWriter(ms))
            {
                writer.AddSheet("Sheet1");
                writer.WriteHeader(XlsxIO_TestSupport.MakeCols());
                writer.WriteRows(new[] { new XlsxIO_TestSupport.Order { A = 1 } }, XlsxIO_TestSupport.MakeTypedPlan(new[] { "A" }));
                writer.AddTable(new TableDefinition("T1", "A1:A2", showTotalsRow: true).WithColumn("A"));
            }
            using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
            zip.GetEntry("xl/tables/table1_1.xml").ShouldNotBeNull();
            var xml = XlsxIO_TestSupport.ReadEntry(ms.ToArray(), "xl/tables/table1_1.xml");
            xml.ShouldContain("totalsRowShown=\"1\"");
            xml.ShouldNotContain("<totalsRowShown");
        }

        [Fact]
        public void AddTable_InvalidRef_Throws()
        {
            using var ms = new MemoryStream();
            using var writer = new XlsxWriter(ms);
            writer.AddSheet("Sheet1");

            Should.Throw<ArgumentException>(() => writer.AddTable(new TableDefinition("T1", "A1:&bad")));
        }

        [Fact]
        public void AddTable_ValidatesColumnCountAndNames()
        {
            using var ms = new MemoryStream();
            using var writer = new XlsxWriter(ms);
            writer.AddSheet("Sheet1");
            Should.Throw<ArgumentException>(() => writer.AddTable(new TableDefinition("T1", "A1:B2").WithColumn("A")));
            Should.Throw<ArgumentException>(() => writer.AddTable(new TableDefinition("T2", "A1:B2").WithColumn("A").WithColumn("A")));
        }

        [Fact]
        public void Complete_ValidatesTableNamesAcrossSheets()
        {
            using var ms = new MemoryStream();
            using var writer = new XlsxWriter(ms);
            writer.AddSheet("First");
            writer.AddTable(new TableDefinition("T1", "A1:A2").WithColumn("A"));
            writer.AddSheet("Second");
            writer.AddTable(new TableDefinition("T1", "A1:A2").WithColumn("A"));
            Should.Throw<ArgumentException>(() => writer.Complete());
        }

        [Fact]
        public async Task Write_RelaxedCellReferences_OmitsCellRefsAndRemainsReadable()
        {
            var ascii = Enumerable.Range(0, 100).Select(i => new StringDto { Name = $"name-{i}" }).ToList();
            var bytes = Xlsx.ToBytes(ascii, options: new XlsxWriteOptions { StrictCellReferences = false });
            bytes.ShouldNotBeEmpty();

            string sheetXml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            sheetXml.ShouldNotContain(" r=\"A1\"");
            sheetXml.ShouldContain("<c t=\"inlineStr\"><is><t>Name</t></is></c>");

            var roundtrip = Xlsx.Read<StringDto>(new MemoryStream(bytes)).ToList();
            roundtrip.Count.ShouldBe(100);
            roundtrip[0].Name.ShouldBe("name-0");
            roundtrip[^1].Name.ShouldBe("name-99");
        }
    }
}
