using GenFu;
using Magicodes.ExporterAndImporter.Pdf;
using Magicodes.ExporterAndImporter.Tests.Models;
using Shouldly;
using System.IO;
using System.Threading.Tasks;
using Xunit;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class PdfExporter_Tests
    {
        [Fact(DisplayName = "导出Pdf测试")]
        public async Task ExportPdf_Test()
        {
            var exporter = new PdfExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(ExportPdf_Test) + ".pdf");
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            //此处使用默认模板导出
            var result = await exporter.ExportByTemplate(filePath, A.ListOf<ExportTestData>());
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
        }


        [Fact(DisplayName = "自定义模板导出Pdf测试")]
        public async Task ExportPdfByTemplate_Test()
        {
            var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "ExportTemplates", "tpl1.cshtml");
            var tpl = File.ReadAllText(tplPath);
            var exporter = new PdfExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(ExportPdfByTemplate_Test) + ".pdf");
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
