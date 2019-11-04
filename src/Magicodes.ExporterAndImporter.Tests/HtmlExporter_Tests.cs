// ======================================================================
// 
//           Copyright (C) 2019-2030 湖南心莱信息科技有限公司
//           All rights reserved
// 
//           filename : HtmlExporter_Tests.cs
//           description :
// 
//           created by 雪雁 at  2019-09-26 14:59
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using System;
using System.IO;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Html;
using Magicodes.ExporterAndImporter.Tests.Models;
using Magicodes.ExporterAndImporter.Tests.Models.Export;
using Shouldly;
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
            if (File.Exists(filePath)) File.Delete(filePath);
            //此处使用默认模板导出
            var result = await exporter.ExportListByTemplate(filePath, GenFu.GenFu.ListOf<ExportTestData>());
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
            if (File.Exists(filePath)) File.Delete(filePath);
            //此处使用默认模板导出
            var result = await exporter.ExportListByTemplate(filePath,
                GenFu.GenFu.ListOf<ExportTestData>(), tpl);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
        }

        [Fact(DisplayName = "导出收据")]
        public async Task ExportReceipt_Test()
        {
            var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "ExportTemplates", "receipt.cshtml");
            var tpl = File.ReadAllText(tplPath);
            var exporter = new HtmlExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(ExportReceipt_Test) + ".html");
            if (File.Exists(filePath)) File.Delete(filePath);
            //此处使用默认模板导出
            var result = await exporter.ExportByTemplate(filePath,
                new ReceiptInfo()
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
    }
}