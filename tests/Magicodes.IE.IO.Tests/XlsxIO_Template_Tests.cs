
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
        public async Task Template_ReplacesSingleVar()
        {
            var templatePath = Path.Combine(Path.GetTempPath(), $"tpl_single_{Guid.NewGuid():N}.xlsx");
            var outputPath = Path.Combine(Path.GetTempPath(), $"out_single_{Guid.NewGuid():N}.xlsx");
            try
            {
                using (var fs = File.Create(templatePath))
                using (var zip = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Create))
                {
                    var entry = zip.CreateEntry("xl/worksheets/sheet1.xml");
                    using var es = entry.Open();
                    using var sw = new StreamWriter(es);
                    sw.Write(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><sheetData><row r=""1""><c r=""A1"" t=""inlineStr""><is><t>客户:{{CustomerName}}</t></is></c></row><row r=""2""><c r=""A2"" t=""inlineStr""><is><t>订单号:{{OrderNo}}</t></is></c></row></sheetData></worksheet>");
                }
                var data = new TemplateHolder { CustomerName = "张三", OrderNo = "SO-1" };
                await Xlsx.ExportByTemplateAsync(templatePath, outputPath, data);
                File.Exists(outputPath).ShouldBeTrue();
                using var za = new ZipArchive(File.OpenRead(outputPath), ZipArchiveMode.Read);
                using var sr = new StreamReader(za.GetEntry("xl/worksheets/sheet1.xml")!.Open());
                var xml = await sr.ReadToEndAsync();
                xml.ShouldContain("客户:张三");
                xml.ShouldContain("订单号:SO-1");
            }
            finally
            {
                if (File.Exists(templatePath)) File.Delete(templatePath);
                if (File.Exists(outputPath)) File.Delete(outputPath);
            }
        }

        [Fact]
        public async Task Template_EscapesXmlTextValues()
        {
            using var templateMs = new MemoryStream();
            using (var zip = new ZipArchive(templateMs, ZipArchiveMode.Create, leaveOpen: true))
            {
                var entry = zip.CreateEntry("xl/worksheets/sheet1.xml");
                using var es = entry.Open();
                using var sw = new StreamWriter(es);
                sw.Write(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><sheetData><row r=""1""><c r=""A1"" t=""inlineStr""><is><t>{{CustomerName}}</t></is></c></row></sheetData></worksheet>");
            }

            templateMs.Position = 0;
            using var outputMs = new MemoryStream();
            await Xlsx.ExportByTemplateAsync(templateMs, outputMs, new TemplateHolder { CustomerName = "A&B<C>" });

            var xml = XlsxIO_TestSupport.ReadEntry(outputMs.ToArray(), "xl/worksheets/sheet1.xml");
            xml.ShouldContain("A&amp;B&lt;C&gt;");
            XlsxIO_TestSupport.AssertWellFormedXml(xml, "sheet1.xml");
        }

        [Fact]
        public async Task Template_ReplacesListBlock()
        {
            var templatePath = Path.Combine(Path.GetTempPath(), $"tpl_list_{Guid.NewGuid():N}.xlsx");
            var outputPath = Path.Combine(Path.GetTempPath(), $"out_list_{Guid.NewGuid():N}.xlsx");
            try
            {
                using (var fs = File.Create(templatePath))
                using (var zip = new System.IO.Compression.ZipArchive(fs, System.IO.Compression.ZipArchiveMode.Create))
                {
                    var entry = zip.CreateEntry("xl/worksheets/sheet1.xml");
                    using var es = entry.Open();
                    using var sw = new StreamWriter(es);
                    sw.Write(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><sheetData><row r=""1""><c r=""A1"" t=""inlineStr""><is><t>客户:{{CustomerName}}</t></is></c></row>{{#Items}}<row r=""1""><c r=""A1"" t=""inlineStr""><is><t>{{Name}} x{{Qty}} ${{Price}}</t></is></c></row>{{/Items}}</sheetData></worksheet>");
                }
                var data = new TemplateHolder
                {
                    CustomerName = "李四",
                    Items = new() { new TemplateCell { Name = "A", Qty = 2, Price = 5m }, new TemplateCell { Name = "B", Qty = 3, Price = 7m } }
                };
                await Xlsx.ExportByTemplateAsync(templatePath, outputPath, data);
                using var za = new ZipArchive(File.OpenRead(outputPath), ZipArchiveMode.Read);
                using var sr = new StreamReader(za.GetEntry("xl/worksheets/sheet1.xml")!.Open());
                var xml = await sr.ReadToEndAsync();
                xml.ShouldContain("客户:李四");
                xml.ShouldContain("A x2 $5");
                xml.ShouldContain("B x3 $7");
                xml.ShouldContain("<row r=\"2\">");
                xml.ShouldContain("<row r=\"3\">");
                xml.ShouldContain("r=\"A2\"");
                xml.ShouldContain("r=\"A3\"");
            }
            finally
            {
                if (File.Exists(templatePath)) File.Delete(templatePath);
                if (File.Exists(outputPath)) File.Delete(outputPath);
            }
        }

        [Fact]
        public async Task Template_ReplacesSharedStringsPlaceholders()
        {
            using var templateMs = new MemoryStream();
            using (var zip = new ZipArchive(templateMs, ZipArchiveMode.Create, leaveOpen: true))
            {
                var sst = zip.CreateEntry("xl/sharedStrings.xml");
                using (var es = sst.Open())
                using (var sw = new StreamWriter(es))
                    sw.Write(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><sst xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"" count=""1"" uniqueCount=""1""><si><t>客户:{{CustomerName}}</t></si></sst>");

                var sheet = zip.CreateEntry("xl/worksheets/sheet1.xml");
                using (var es = sheet.Open())
                using (var sw = new StreamWriter(es))
                    sw.Write(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><sheetData><row r=""1""><c r=""A1"" t=""s""><v>0</v></c></row></sheetData></worksheet>");
            }

            templateMs.Position = 0;
            using var outputMs = new MemoryStream();
            await Xlsx.ExportByTemplateAsync(templateMs, outputMs, new TemplateHolder { CustomerName = "赵六" });

            var sstXml = XlsxIO_TestSupport.ReadEntry(outputMs.ToArray(), "xl/sharedStrings.xml");
            sstXml.ShouldContain("客户:赵六");
            sstXml.ShouldNotContain("{{CustomerName}}");
        }

        [Fact]
        public async Task Template_ReplacesAllWorksheetParts()
        {
            using var templateMs = new MemoryStream();
            using (var zip = new ZipArchive(templateMs, ZipArchiveMode.Create, leaveOpen: true))
            {
                var sheet1 = zip.CreateEntry("xl/worksheets/sheet1.xml");
                using (var es = sheet1.Open())
                using (var sw = new StreamWriter(es))
                    sw.Write(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><sheetData><row r=""1""><c r=""A1"" t=""inlineStr""><is><t>{{CustomerName}}</t></is></c></row></sheetData></worksheet>");

                var sheet2 = zip.CreateEntry("xl/worksheets/sheet2.xml");
                using (var es = sheet2.Open())
                using (var sw = new StreamWriter(es))
                    sw.Write(@"<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?><worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""><sheetData><row r=""1""><c r=""A1"" t=""inlineStr""><is><t>{{OrderNo}}</t></is></c></row></sheetData></worksheet>");
            }

            templateMs.Position = 0;
            using var outputMs = new MemoryStream();
            await Xlsx.ExportByTemplateAsync(templateMs, outputMs, new TemplateHolder { CustomerName = "孙七", OrderNo = "SO-7" });

            var sheet1Xml = XlsxIO_TestSupport.ReadEntry(outputMs.ToArray(), "xl/worksheets/sheet1.xml");
            var sheet2Xml = XlsxIO_TestSupport.ReadEntry(outputMs.ToArray(), "xl/worksheets/sheet2.xml");

            sheet1Xml.ShouldContain("孙七");
            sheet2Xml.ShouldContain("SO-7");
        }
    }
}
