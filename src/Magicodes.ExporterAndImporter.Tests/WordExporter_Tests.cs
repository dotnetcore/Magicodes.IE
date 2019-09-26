using GenFu;
using Magicodes.ExporterAndImporter.Tests.Models;
using Magicodes.ExporterAndImporter.Word;
using Shouldly;
using System.IO;
using System.Threading.Tasks;
using Xunit;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class WordExporter_Tests
    {
        [Fact(DisplayName = "导出Word测试")]
        public async Task ExportWord_Test()
        {
            var exporter = new WordExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(ExportWord_Test) + ".docx");
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
            //此处使用默认模板导出
            var result = await exporter.ExportByTemplate(filePath, A.ListOf<ExportTestData>());
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
        }


        [Fact(DisplayName = "自定义模板导出Word测试")]
        public async Task ExportWordByTemplate_Test()
        {
            var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "ExportTemplates", "tpl1.cshtml");
            var tpl = File.ReadAllText(tplPath);
            var exporter = new WordExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(ExportWordByTemplate_Test) + ".docx");
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
