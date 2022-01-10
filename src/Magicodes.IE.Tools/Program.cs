using System;
using System.Reflection;
using System.Linq;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using System.Collections.Generic;
using System.IO;

namespace Magicodes.IE.Tools
{
    internal class Program
    {
        private static void Main(string[] args)
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
                IExporter exporter = new ExcelExporter();
                var data = new List<ExportTestDataWithPicture>
                {
                    new ExportTestDataWithPicture
                    {
                        Img = "https://gitee.com/magicodes/Magicodes.IE/raw/master/docs/Magicodes.IE.png"
                    },
                    new ExportTestDataWithPicture
                    {
                        Img = "https://gitee.com/magicodes/Magicodes.IE/raw/master/res/wechat.jpg"
                    }
                };

                var filePath = Path.Combine(System.IO.Directory.GetCurrentDirectory(), "test.xlsx");
                var result = exporter.Export("test.xlsx", data).Result;
                Console.WriteLine($"导出成功：{filePath}！");
            }
        }

        [ExcelExporter(Name = "测试")]
        public class ExportTestDataWithPicture
        {
            [ExportImageField(Width = 50, Height = 120, Alt = "404")]
            [ExporterHeader(DisplayName = "图", IsAutoFit = false)]
            public string Img { get; set; }
        }
    }
}