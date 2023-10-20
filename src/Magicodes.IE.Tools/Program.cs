using System;
using System.Reflection;
using System.Linq;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using System.Collections.Generic;
using System.IO;
using Magicodes.Benchmarks.Models;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Excel.Utility;
using System.Dynamic;

namespace Magicodes.IE.Tools
{
    internal class Program
    {
        private readonly static List<ExportTestDataWithAttrs> _exportTestData = new List<ExportTestDataWithAttrs>();
        private static async Task Main(string[] args)
        {
            if (args.Length == 0)
            {
                var versionString = Assembly.GetEntryAssembly()
                                        .GetCustomAttribute<AssemblyInformationalVersionAttribute>()
                                        .InformationalVersion
                                        .ToString();

                Console.WriteLine($"mie v{versionString}");
                Console.WriteLine("-------------");
                Console.WriteLine("\nGithub:");
                Console.WriteLine("  https://github.com/dotnetcore/Magicodes.IE");
                return;
            }
            else if (args.Any(p => "TEST".Equals(p, StringComparison.CurrentCultureIgnoreCase)))
            {
                ExcelExporter exporter = new ExcelExporter();
                var data = new List<ExportTestDataWithPicture>
                {
                    new ExportTestDataWithPicture
                    {
                        Img = Path.Combine(Directory.GetCurrentDirectory(), "zero-DPI.Jpeg"),
                        Text ="张三"
                    }
                };

                Parallel.For(0, 10, (i) =>
                {
                    data.Add(new ExportTestDataWithPicture
                    {
                        Img = Path.Combine(Directory.GetCurrentDirectory(), "zero-DPI.Jpeg"),
                        Text = "张三"
                    });
                });

                var filePath = Path.Combine(System.IO.Directory.GetCurrentDirectory(), "test.xlsx");
                var result = await exporter.Export("test.xlsx", data);
                Console.WriteLine($"导出成功：{filePath}！");
            }
            Console.WriteLine("完成");
            Console.ReadLine();
        }

        static async Task Test()
        {
            for (var i = 1; i <= 10000000; i++)
            {
                _exportTestData.Add(new ExportTestDataWithAttrs
                {
                    Age = i,
                    Name = "Mr.A",
                    Text3 = "Text3"
                });
            }
            var helper = new ExportHelper<ExportTestDataWithAttrs>();

            using (helper.CurrentExcelPackage)
            {
                var sheetCount = (int)(_exportTestData.Count / helper.ExcelExporterSettings.MaxRowNumberOnASheet) +
                                 ((_exportTestData.Count % helper.ExcelExporterSettings.MaxRowNumberOnASheet) > 0
                                     ? 1
                                     : 0);
                for (int i = 0; i < sheetCount; i++)
                {
                    var sheetDataItems = _exportTestData.Skip(i * helper.ExcelExporterSettings.MaxRowNumberOnASheet)
                        .Take(helper.ExcelExporterSettings.MaxRowNumberOnASheet).ToList();
                    helper.AddExcelWorksheet();
                    helper.Export(sheetDataItems);
                }
                await helper.CurrentExcelPackage.GetAsByteArrayAsync();
            }
        }

        [ExcelExporter(Name = "测试")]
        public class ExportTestDataWithPicture
        {
            [ExportImageField(Width = 50, Height = 120, Alt = "404")]
            [ExporterHeader(DisplayName = "图", IsAutoFit = false)]
            public string Img { get; set; }

            public string Text { get; set; }
        }
    }
}