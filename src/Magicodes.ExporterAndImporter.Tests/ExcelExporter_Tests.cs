// ======================================================================
// 
//           Copyright (C) 2019-2030 湖南心莱信息科技有限公司
//           All rights reserved
// 
//           filename : ExcelExporter_Tests.cs
//           description :
// 
//           created by 雪雁 at  2019-09-11 13:51
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Excel.Builder;
using Magicodes.ExporterAndImporter.Pdf;
using Magicodes.ExporterAndImporter.Tests.Models;
using Shouldly;
using Xunit;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class ExcelExporter_Tests
    {
        [Fact(DisplayName = "特性导出")]
        public async Task AttrsExport_Test()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "testAttrs.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var result = await exporter.Export(filePath, new List<ExportTestDataWithAttrs>
            {
                new ExportTestDataWithAttrs
                {
                    Text = "啊实打实大苏打撒",
                    Name = "aa",
                    Number = 5000,
                    Text2 = "w萨达萨达萨达撒",
                    Text3 = "sadsad打发打发士大夫的"
                },
                new ExportTestDataWithAttrs
                {
                    Text = "啊实打实大苏打撒",
                    Name = "啊实打实大苏打撒",
                    Number = 6000,
                    Text2 = "w萨达萨达萨达撒",
                    Text3 = "sadsad打发打发士大夫的"
                },
                new ExportTestDataWithAttrs
                {
                    Text = "啊实打实速度大苏打撒",
                    Name = "萨达萨达",
                    Number = 6000,
                    Text2 = "突然他也让他人",
                    Text3 = "sadsad打发打发士大夫的"
                }
            });
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
        }

        [Fact(DisplayName = "多语言特性导出")]
        public async Task AttrsLocalizationExport_Test()
        {
            IExporter exporter = new ExcelExporter();
            ExcelBuilder.Create().WithColumnHeaderStringFunc(key =>
            {
                if (key.Contains("文本")) return "Text";
                return "未知语言";
            }).Build();

            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "testAttrsLocalization.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var result = await exporter.Export(filePath, new List<AttrsLocalizationTestData>
            {
                new AttrsLocalizationTestData
                {
                    Text = "啊实打实大苏打撒",
                    Name = "aa",
                    Number = 5000,
                    Text2 = "w萨达萨达萨达撒",
                    Text3 = "sadsad打发打发士大夫的"
                },
                new AttrsLocalizationTestData
                {
                    Text = "啊实打实大苏打撒",
                    Name = "啊实打实大苏打撒",
                    Number = 6000,
                    Text2 = "w萨达萨达萨达撒",
                    Text3 = "sadsad打发打发士大夫的"
                },
                new AttrsLocalizationTestData
                {
                    Text = "啊实打实速度大苏打撒",
                    Name = "萨达萨达",
                    Number = 6000,
                    Text2 = "突然他也让他人",
                    Text3 = "sadsad打发打发士大夫的"
                }
            });
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
        }

        [Fact(DisplayName = "导出Excel")]
        public async Task Export_Test()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "test.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var result = await exporter.Export(filePath, new List<ExportTestData>
            {
                new ExportTestData
                {
                    Name1 = "1",
                    Name2 = "test",
                    Name3 = "12",
                    Name4 = "11"
                },
                new ExportTestData
                {
                    Name1 = "1",
                    Name2 = "test",
                    Name3 = "12",
                    Name4 = "11"
                }
            });
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
        }


        [Fact(DisplayName = "导出PDF测试")]
        public async Task ExportPDF_Test()
        {
            var exporter = new PdfExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "test.pdf");
            if (File.Exists(filePath)) File.Delete(filePath);
            var result = await exporter.ExportByTemplate(filePath, new List<ExportTestData>
            {
                new ExportTestData
                {
                    Name1 = "1",
                    Name2 = "test",
                    Name3 = "12",
                    Name4 = "11"
                },
                new ExportTestData
                {
                    Name1 = "1",
                    Name2 = "test",
                    Name3 = "12",
                    Name4 = "11"
                }
            });
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
        }
    }
}