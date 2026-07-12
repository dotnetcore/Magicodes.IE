
using System;
using System.IO;
using System.IO.Compression;
using System.Linq;
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
        public async Task Style_Bold_WritesFontInStyles()
        {
            var p = new ExportProfile<OrderDto>().Column(x => x.OrderNo, c => c.WithBold());
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/styles.xml");
            xml.ShouldContain("<b/>");
        }

        [Fact]
        public async Task Style_BackgroundColor_WritesSolidFill()
        {
            var p = new ExportProfile<OrderDto>().Column(x => x.OrderNo, c => c.WithBackgroundColor("FFFF00"));
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/styles.xml");
            xml.ShouldContain("<patternFill patternType=\"solid\"");
            xml.ShouldContain("FFFF00");
        }

        [Fact]
        public async Task Style_FontSize_WritesCustomSize()
        {
            var p = new ExportProfile<OrderDto>().Column(x => x.OrderNo, c => c.WithFontSize(20));
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/styles.xml");
            xml.ShouldContain("val=\"20\"");
        }

        [Fact]
        public async Task Style_FontColor_WritesColor()
        {
            var p = new ExportProfile<OrderDto>().Column(x => x.OrderNo, c => c.WithFontColor("FF0000"));
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/styles.xml");
            xml.ShouldContain("FF0000");
        }

        [Fact]
        public async Task Style_FontName_WritesCustomName()
        {
            var p = new ExportProfile<OrderDto>().Column(x => x.OrderNo, c => c.WithFontName("Consolas"));
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/styles.xml");
            xml.ShouldContain("Consolas");
        }

        [Fact]
        public async Task Style_Border_WritesThinBlackBorder()
        {
            var p = new ExportProfile<OrderDto>().Column(x => x.OrderNo, c => c.WithBorderStyle(BorderStyle.Thin));
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/styles.xml");
            xml.ShouldContain("<border>");
            xml.ShouldContain("style=\"thin\"");
        }

        [Fact]
        public async Task Style_Border_CustomColor()
        {
            var p = new ExportProfile<OrderDto>().Column(x => x.OrderNo, c => c.WithBorderStyle(BorderStyle.Thin, "FF0000"));
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/styles.xml");
            xml.ShouldContain("FF0000");
        }

        [Fact]
        public async Task Style_Border_Dashed_WritesDashedBorder()
        {
            var p = new ExportProfile<OrderDto>().Column(x => x.OrderNo, c => c.WithBorderStyle(BorderStyle.Dashed));
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/styles.xml");
            xml.ShouldContain("style=\"dashed\"");
        }

        [Fact]
        public async Task Style_Border_Dotted_WritesDottedBorder()
        {
            var p = new ExportProfile<OrderDto>().Column(x => x.OrderNo, c => c.WithBorderStyle(BorderStyle.Dotted));
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/styles.xml");
            xml.ShouldContain("style=\"dotted\"");
        }

        [Fact]
        public async Task Style_Border_Double_WritesDoubleBorder()
        {
            var p = new ExportProfile<OrderDto>().Column(x => x.OrderNo, c => c.WithBorderStyle(BorderStyle.Double));
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/styles.xml");
            xml.ShouldContain("style=\"double\"");
        }

        [Fact]
        public async Task Style_Italic_WritesItalic()
        {
            var p = new ExportProfile<OrderDto>().Column(x => x.OrderNo, c => c.WithItalic());
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/styles.xml");
            xml.ShouldContain("<i/>");
        }

        [Fact]
        public async Task Style_Underline_WritesUnderline()
        {
            var p = new ExportProfile<OrderDto>().Column(x => x.OrderNo, c => c.WithUnderline());
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/styles.xml");
                        xml.ShouldContain("<u ");
        }

        [Fact]
        public async Task Style_StrikeThrough_WritesStrikeThrough()
        {
            var p = new ExportProfile<OrderDto>().Column(x => x.OrderNo, c => c.WithStrikeThrough());
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/styles.xml");
            xml.ShouldContain("<strike/>");
        }

        [Fact]
        public async Task Style_VerticalAlignment_Center_WritesCenter()
        {
            var p = new ExportProfile<OrderDto>().Column(x => x.OrderNo, c => c.WithVerticalAlignment(VerticalAlignment.Center));
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/styles.xml");
            xml.ShouldContain("vertical=\"center\"");
        }

        [Fact]
        public async Task Style_VerticalAlignment_Top_WritesTop()
        {
            var p = new ExportProfile<OrderDto>().Column(x => x.OrderNo, c => c.WithVerticalAlignment(VerticalAlignment.Top));
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/styles.xml");
            xml.ShouldContain("vertical=\"top\"");
        }

        [Fact]
        public async Task RowHeight_Default_WritesSheetFormatPr()
        {
            var p = new ExportProfile<OrderDto>().WithDefaultRowHeight(20);
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            xml.ShouldContain("<sheetFormatPr");
            xml.ShouldContain("defaultRowHeight=\"15\"");
        }

        [Fact]
        public async Task Style_Wrap_WritesWrapText()
        {
            var p = new ExportProfile<OrderDto>().Column(x => x.OrderNo, c => c.WithWrap());
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/styles.xml");
            xml.ShouldContain("wrapText=\"1\"");
        }

        [Fact]
        public async Task Style_BoldHeader_RealXfIdOnCell()
        {
            var p = new ExportProfile<OrderDto>().Column(x => x.OrderNo, c => c.WithBold());
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/styles.xml");
            xml.ShouldContain("<b/>");
        }

        [Fact]
        public async Task RowHeight_WritesSheetFormatPr()
        {
            var p = new ExportProfile<OrderDto>().WithDefaultRowHeight(15);
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            xml.ShouldContain("defaultRowHeight=\"11.25\"");
        }

        [Fact]
        public async Task RowHeight_FromColumnConfig_WritesHeaderRowHeight()
        {
            var p = new ExportProfile<OrderDto>()
                .Column(x => x.OrderNo, c => c.WithRowHeight(30));
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            xml.ShouldContain("ht=\"22.5\"");
        }

        [Fact]
        public async Task EveryCell_HasRAttribute_RequiredBySpec()
        {
                        var bytes = Xlsx.ToBytes(new[]
            {
                new OrderDto { OrderNo = "A1", Amount = 1m, CreatedAt = new DateTime(2024, 1, 1) },
                new OrderDto { OrderNo = "A2", Amount = 2m, CreatedAt = new DateTime(2024, 1, 2) },
            });
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            int cellCount = xml.Split(new[] { "<c " }, StringSplitOptions.RemoveEmptyEntries).Length - 1;
            int cellsWithR = System.Text.RegularExpressions.Regex.Matches(xml, "<c r=\"[A-Z]+\\d+\"").Count;
            cellCount.ShouldBe(cellsWithR, $"每个 <c> 都需要 r=\"A1\" — 缺 {cellCount - cellsWithR} 个");
        }

        [Fact]
        public async Task ContentTypes_IncludesGifAndBmp()
        {
                        var bytes = Xlsx.ToBytes(new[] { new OrderDto() });
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "[Content_Types].xml");
            xml.ShouldContain("Extension=\"gif\"");
            xml.ShouldContain("Extension=\"bmp\"");
        }

        [Fact]
        public async Task Worksheet_ElementOrder_SpecCompliant()
        {
            var p = new ExportProfile<OrderDto>()
                .WithDefaultRowHeight(20)
                .Column(x => x.OrderNo, c => c.WithWidth(15));
            var bytes = Xlsx.ToBytes(new[] { new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            int sheetViewsPos = xml.IndexOf("<sheetViews");
            int colsPos = xml.IndexOf("<cols");
            int sheetFormatPos = xml.IndexOf("<sheetFormatPr");
            int sheetDataPos = xml.IndexOf("<sheetData");
            sheetViewsPos.ShouldBeGreaterThan(0);
            sheetFormatPos.ShouldBeGreaterThan(sheetViewsPos);
            colsPos.ShouldBeGreaterThan(sheetFormatPos);
            sheetDataPos.ShouldBeGreaterThan(colsPos);
        }

        [Fact]
        public async Task Worksheet_Closing_ElementOrder_SpecCompliant()
        {
            var p = new ExportProfile<OrderDto>().WithAutoFilter("A1:C2");
            var bytes = Xlsx.ToBytes(new[] { new OrderDto(), new OrderDto() }, p);
            var xml = XlsxIO_TestSupport.ReadEntry(bytes, "xl/worksheets/sheet1.xml");
            int autoFilterPos = xml.IndexOf("<autoFilter");
            int pageMarginsPos = xml.IndexOf("<pageMargins");
            int sheetDataEnd = xml.IndexOf("</sheetData>");
            autoFilterPos.ShouldBeGreaterThan(sheetDataEnd);
            pageMarginsPos.ShouldBeGreaterThan(sheetDataEnd);
        }
    }
}
