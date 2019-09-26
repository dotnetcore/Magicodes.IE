using GenFu;
using Magicodes.ExporterAndImporter.Html;
using Magicodes.ExporterAndImporter.Tests.Models;
using Shouldly;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Xunit;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class HtmlExporter_Tests
    {
        [Fact(DisplayName = "导出Html测试")]
        public async Task ExportHtml_Test()
        {
            var exporter = new HtmlExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(ExportHtml_Test) + ".html");
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            //此处使用默认模板导出
            var result = await exporter.ExportByTemplate(filePath, A.ListOf<ExportTestData>());
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
        }


        [Fact(DisplayName = "自定义模板导出Html测试")]
        public async Task ExportHtmlByTemplate_Test()
        {
            var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "ExportTemplates", "tpl1.cshtml");
            var tpl = File.ReadAllText(tplPath);
            var exporter = new HtmlExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(ExportHtmlByTemplate_Test) + ".html");
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            //此处使用默认模板导出
            var result = await exporter.ExportByTemplate(filePath,
                A.ListOf<ExportTestData>(), tpl);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
        }
    }
}
