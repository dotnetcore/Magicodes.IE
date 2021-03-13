// ======================================================================
//
//           filename : WordExporter_Tests.cs
//           description :
//
//           created by 雪雁 at  2019-09-26 14:59
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
//
// ======================================================================
using Magicodes.ExporterAndImporter.Tests.Models.Export;
using Magicodes.ExporterAndImporter.Word;
using Shouldly;
using System;
using System.IO;
using System.Threading.Tasks;
using Xunit;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class WordExporter_Tests : TestBase
    {
        [Fact(DisplayName = "导出Word测试")]
        public async Task ExportWord_Test()
        {
            var exporter = new WordExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(ExportWord_Test) + ".docx");
            if (File.Exists(filePath)) File.Delete(filePath);
            //此处使用默认模板导出
            var result = await exporter.ExportListByTemplate(filePath, GenFu.GenFu.ListOf<ExportTestData>());
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
        }

        [Fact(DisplayName = "自定义模板导出Word测试")]
        public async Task ExportWordByTemplate_Test()
        {
            var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "ExportTemplates", "tpl1.cshtml");
            var tpl = File.ReadAllText(tplPath);
            var exporter = new WordExporter();
            var ex = await Assert.ThrowsAnyAsync<ArgumentException>(async () => await exporter.ExportListByTemplate(null,
                 GenFu.GenFu.ListOf<ExportTestData>(), tpl));
            ex.Message.ShouldContain("文件名必须填写");

            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(ExportWordByTemplate_Test) + ".docx");
            if (File.Exists(filePath)) File.Delete(filePath);
            //此处使用默认模板导出
            var result = await exporter.ExportListByTemplate(filePath,
                GenFu.GenFu.ListOf<ExportTestData>(), tpl);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
        }

        [Fact(DisplayName = "自定义模板导出WordBytes测试")]
        public async Task ExportWordBytesByTemplate_Test()
        {
            var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "ExportTemplates", "tpl1.cshtml");
            var tpl = File.ReadAllText(tplPath);
            var exporter = new WordExporter();
            var ex = await Assert.ThrowsAnyAsync<ArgumentException>(async () => await exporter.ExportListByTemplate(null,
                 GenFu.GenFu.ListOf<ExportTestData>(), tpl));
            ex.Message.ShouldContain("文件名必须填写");

            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(ExportWordBytesByTemplate_Test) + ".docx");
            if (File.Exists(filePath)) File.Delete(filePath);
            //此处使用默认模板导出
            var result = await exporter.ExportListBytesByTemplate(
                GenFu.GenFu.ListOf<ExportTestData>(), tpl);

            result.ShouldNotBeNull();
            result.Length.ShouldBeGreaterThan(0);

            using (var file = File.OpenWrite(filePath))
            {
                file.Write(result, 0, result.Length);
            }
            File.Exists(filePath).ShouldBeTrue();
        }

        [Fact(DisplayName = "自定义模板导出Word文件测试")]
        public async Task ExportWordFileByTemplate_Test()
        {
            var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "ExportTemplates", "receipt.cshtml");
            var tpl = File.ReadAllText(tplPath);
            var exporter = new WordExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(ExportWordFileByTemplate_Test) + ".docx");
            if (File.Exists(filePath)) File.Delete(filePath);
            //此处使用默认模板导出
            var result = await exporter.ExportByTemplate(filePath,
                new ReceiptInfo
                {
                    Amount = 22939.43M,
                    Grade = "2019秋",
                    IdNo = "43062619890622xxxx",
                    Name = "张三",
                    Payee = "湖南心莱信息科技有限公司",
                    PaymentMethod = "微信支付",
                    Profession = "运动训练",
                    Remark = "学费",
                    TradeStatus = "已完成",
                    TradeTime = DateTime.Now,
                    UppercaseAmount = "贰万贰仟玖佰叁拾玖圆肆角叁分",
                    Code = "19071800001"
                }, tpl);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
        }

        [Fact(DisplayName = "自定义模板导出Word文件测试Type类型测试")]
        public async Task ExportWordFileByTemplate_Type_Test()
        {
            var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "ExportTemplates", "receipt.cshtml");
            var tpl = File.ReadAllText(tplPath);
            var exporter = new WordExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(ExportWordFileByTemplate_Type_Test) + ".docx");
            if (File.Exists(filePath)) File.Delete(filePath);
            //此处使用默认模板导出
            var result = await exporter.ExportBytesByTemplate(
                new ReceiptInfo
                {
                    Amount = 22939.43M,
                    Grade = "2019秋",
                    IdNo = "43062619890622xxxx",
                    Name = "张三",
                    Payee = "湖南心莱信息科技有限公司",
                    PaymentMethod = "微信支付",
                    Profession = "运动训练",
                    Remark = "学费",
                    TradeStatus = "已完成",
                    TradeTime = DateTime.Now,
                    UppercaseAmount = "贰万贰仟玖佰叁拾玖圆肆角叁分",
                    Code = "19071800001"
                }, tpl, typeof(ReceiptInfo));
            result.ShouldNotBeNull();
            result.Length.ShouldBeGreaterThan(0);
            using (var file = File.OpenWrite(filePath))
            {
                file.Write(result, 0, result.Length);
            }
            File.Exists(filePath).ShouldBeTrue();
        }

    }
}