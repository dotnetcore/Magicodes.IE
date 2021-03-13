using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Tests.Models.Export;
using Magicodes.ExporterAndImporter.Tests.Models.Export.ExportByTemplate_Test1;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using Shouldly;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class ExcelTemplateExporter_Tests : TestBase
    {
        #region 模板导出

        [Fact(DisplayName = "Excel模板导出教材订购明细样表（含图片）")]
        public async Task ExportByTemplate_Test()
        {
            //模板路径
            var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "ExportTemplates",
                "2020年春季教材订购明细样表.xlsx");
            //创建Excel导出对象
            IExportFileByTemplate exporter = new ExcelExporter();
            //导出路径
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(ExportByTemplate_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);
            //根据模板导出
            await exporter.ExportByTemplate(filePath,
                new TextbookOrderInfo("湖南心莱信息科技有限公司", "湖南长沙岳麓区", "雪雁", "1367197xxxx", null,
                    DateTime.Now.ToLongDateString(), "https://docs.microsoft.com/en-us/media/microsoft-logo-dark.png",
                    new List<BookInfo>()
                    {
                        new BookInfo(1, "0000000001", "《XX从入门到放弃》", null, "机械工业出版社", "3.14", 100, "备注")
                        {
                            Cover = Path.Combine("TestFiles", "ExporterTest.png")
                        },
                        new BookInfo(2, "0000000001", "《XX从入门到放弃》", null, "机械工业出版社", "3.14", 100, "备注")
                        {
                            Cover = "https://docs.microsoft.com/en-us/media/microsoft-logo-dark.png"
                        },
                        new BookInfo(3, "0000000002", "《XX从入门到放弃》", "张三", "机械工业出版社", "3.14", 100, null),
                        new BookInfo(4, null, "《XX从入门到放弃》", "张三", "机械工业出版社", "3.14", 100, "备注")
                        {
                            Cover = Path.Combine("TestFiles", "issue131.png")
                        }
                    }),
                tplPath);

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //检查转换结果
                var sheet = pck.Workbook.Worksheets.First();
                //确保所有的转换均已完成
                sheet.Cells[sheet.Dimension.Address].Any(p => p.Text.Contains("{{")).ShouldBeFalse();
                //检查图片
                sheet.Drawings.Count.ShouldBe(4);

                sheet.Cells[sheet.Dimension.Address].Any(p => p.Text.Contains("图")).ShouldBeTrue();
                //检查合计是否正确

                sheet.Cells["H11"].Formula.ShouldBe("=SUM(G4:G6,G4)");
                sheet.Cells["H12"].Formula.ShouldBe("=AVERAGE(G4:G6)");
            }
        }

        [Fact(DisplayName = "模板导出大量数据测试")]
        public async Task Export10000ByTemplate_Test()
        {
            //模板路径
            var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "ExportTemplates",
                "Export10000ByTemplate_Test.xlsx");
            //创建Excel导出对象
            IExportFileByTemplate exporter = new ExcelExporter();
            //导出路径
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(Export10000ByTemplate_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var books = GenFu.GenFu.ListOf<BookInfo>(10000);
            //根据模板导出
            await exporter.ExportByTemplate(filePath,
                new TextbookOrderInfo("湖南心莱信息科技有限公司", "湖南长沙岳麓区", "雪雁", "1367197xxxx", null,
                    DateTime.Now.ToLongDateString(), "https://docs.microsoft.com/en-us/media/microsoft-logo-dark.png",
                    books),
                tplPath);
        }

        [Fact(DisplayName = "模板导出动态导出测试")]
        public async Task DynamicExportWithJObjectByTemplate_Test()
        {
            string json = @"{
              'Company': '雪雁',
              'Address': '湖南长沙',
              'Contact': '雪雁',
              'Tel': '136xxx',
              'BookInfos': [
                {'No':'a1','RowNo':1,'Name':'Docker+Kubernetes应用开发与快速上云','EditorInChief':'李文强','PublishingHouse':'机械工业出版社','Price':65,'PurchaseQuantity':10000,'Cover':'https://img9.doubanio.com/view/ark_article_cover/retina/public/135025435.jpg?v=1585121965','Remark':'备注'},
                {'No':'a2','RowNo':2,'Name':'Docker+Kubernetes应用开发与快速上云','EditorInChief':'李文强','PublishingHouse':'机械工业出版社','Price':65,'PurchaseQuantity':10000,'Cover':'https://img9.doubanio.com/view/ark_article_cover/retina/public/135025435.jpg?v=1585121965','Remark':'备注'}
              ]
            }";
            var jobj = JObject.Parse(json);
            //模板路径
            var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "ExportTemplates",
                "DynamicExportTpl.xlsx");
            //创建Excel导出对象
            IExportFileByTemplate exporter = new ExcelExporter();
            //导出路径
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(DynamicExportWithJObjectByTemplate_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            //根据模板导出
            await exporter.ExportByTemplate(filePath, jobj, tplPath);

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //检查转换结果
                var sheet = pck.Workbook.Worksheets.First();
                //确保所有的转换均已完成
                sheet.Cells[sheet.Dimension.Address].Any(p => p.Text.Contains("{{")).ShouldBeFalse();
            }
        }

        [Fact(DisplayName = "模板导出字典动态导出测试")]
        public async Task DynamicExportWithDictionaryByTemplate_Test()
        {
            var data = new Dictionary<string, object>()
            {
                { "Company","雪雁" },
                { "Address", "湖南长沙" },
                { "Contact", "雪雁" },
                { "Tel", "136xxx" },
                { "BookInfos",new List<Dictionary<string,object>>()
                    {
                        new Dictionary<string, object>()
                        {
                            {"No","A1" },
                            {"RowNo",1 },
                            {"Name","Docker+Kubernetes应用开发与快速上云" },
                            {"EditorInChief","李文强" },
                            {"PublishingHouse","机械工业出版社" },
                            {"Price",65 },
                            {"PurchaseQuantity",50000 },
                            {"Cover","https://img9.doubanio.com/view/ark_article_cover/retina/public/135025435.jpg?v=1585121965" },
                            {"Remark","买起" }
                        },
                        new Dictionary<string, object>()
                        {
                            {"No","A2" },
                            {"RowNo",2 },
                            {"Name","Docker+Kubernetes应用开发与快速上云" },
                            {"EditorInChief","李文强" },
                            {"PublishingHouse","机械工业出版社" },
                            {"Price",65 },
                            {"PurchaseQuantity",50000 },
                            {"Cover","https://img9.doubanio.com/view/ark_article_cover/retina/public/135025435.jpg?v=1585121965" },
                            {"Remark","k8s真香" }
                        }
                    }
                }
            };
            //模板路径
            var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "ExportTemplates",
                "DynamicExportTpl.xlsx");
            //创建Excel导出对象
            IExportFileByTemplate exporter = new ExcelExporter();
            //导出路径
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(DynamicExportWithDictionaryByTemplate_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            //根据模板导出
            await exporter.ExportByTemplate(filePath, data, tplPath);

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //检查转换结果
                var sheet = pck.Workbook.Worksheets.First();
                //确保所有的转换均已完成
                sheet.Cells[sheet.Dimension.Address].Any(p => p.Text.Contains("{{")).ShouldBeFalse();
            }
        }

        [Fact(DisplayName = "模板导出之ExpandoObject动态导出测试")]
        public async Task DynamicExportWithExpandoObjectByTemplate_Test()
        {
            dynamic data = new ExpandoObject();
            data.Company = "雪雁";
            data.Address = "湖南长沙";
            data.Contact = "雪雁";
            data.Tel = "136xxx";
            data.BookInfos = new List<ExpandoObject>() { };

            dynamic book1 = new ExpandoObject();
            book1.No = "A1";
            book1.RowNo = 1;
            book1.Name = "Docker+Kubernetes应用开发与快速上云";
            book1.EditorInChief = "李文强";
            book1.PublishingHouse = "机械工业出版社";
            book1.Price = 65;
            book1.PurchaseQuantity = 50000;
            book1.Cover = "https://img9.doubanio.com/view/ark_article_cover/retina/public/135025435.jpg?v=1585121965";
            book1.Remark = "买买买";
            data.BookInfos.Add(book1);

            dynamic book2 = new ExpandoObject();
            book2.No = "A2";
            book2.RowNo = 2;
            book2.Name = "Docker+Kubernetes应用开发与快速上云";
            book2.EditorInChief = "李文强";
            book2.PublishingHouse = "机械工业出版社";
            book2.Price = 65;
            book2.PurchaseQuantity = 50000;
            book2.Cover = "https://img9.doubanio.com/view/ark_article_cover/retina/public/135025435.jpg?v=1585121965";
            book2.Remark = "买买买";
            data.BookInfos.Add(book2);

            //模板路径
            var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "ExportTemplates",
                "DynamicExportTpl.xlsx");
            //创建Excel导出对象
            IExportFileByTemplate exporter = new ExcelExporter();
            //导出路径
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(DynamicExportWithExpandoObjectByTemplate_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            //根据模板导出
            await exporter.ExportByTemplate(filePath, data, tplPath);

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //检查转换结果
                var sheet = pck.Workbook.Worksheets.First();
                //确保所有的转换均已完成
                sheet.Cells[sheet.Dimension.Address].Any(p => p.Text.Contains("{{")).ShouldBeFalse();
            }

        }

        /// <summary>
        /// https://github.com/dotnetcore/Magicodes.IE/issues/34
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "Excel模板导出测试（issues#34）")]
        public async Task ExportByTemplate_Test1()
        {
            //模板路径
            var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "ExportTemplates",
                "template.xlsx");
            //创建Excel导出对象
            IExportFileByTemplate exporter = new ExcelExporter();
            //导出路径
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(ExportByTemplate_Test1) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var airCompressors = new List<AirCompressor>
            {
                new AirCompressor()
                {
                    Name = "1#",
                    Manufactor = "111",
                    ExhaustPressure = "0",
                    ExhaustTemperature = "66.7-95",
                    RunningTime = "35251",
                    WarningError = "正常",
                    Status = "开机"
                },
                new AirCompressor()
                {
                    Name = "2#",
                    Manufactor = "222",
                    ExhaustPressure = "1",
                    ExhaustTemperature = "90.7-95",
                    RunningTime = "2222",
                    WarningError = "正常",
                    Status = "开机"
                }
            };

            var afterProcessings = new List<AfterProcessing>
            {
                new AfterProcessing()
                {
                    Name = "1#abababa",
                    Manufactor = "杭州立山",
                    RunningTime = "NaN",
                    WarningError = "故障",
                    Status = "停机"
                }
            };

            var suggests = new List<Suggest>
            {
                new Suggest()
                {
                    Number = 1,
                    Description = "故障停机",
                    SuggestMessage = "顾问团队远程协助"
                }
            };

            //根据模板导出
            await exporter.ExportByTemplate(filePath,
                new ReportInformation()
                {
                    Contacts = "11112",
                    ContactsNumber = "13642666666",
                    CustomerName = "ababace",
                    Date = DateTime.Now.ToString("yyyy年MM月dd日"),
                    SystemExhaustPressure = "0.54-0.62",
                    SystemDewPressure = "-0.63--77.5",
                    SystemDayFlow = "201864",
                    AirCompressors = airCompressors,
                    AfterProcessings = afterProcessings,
                    Suggests = suggests,
                    SystemPressureHisotries = new List<SystemPressureHisotry>()
                },
                tplPath);

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //检查转换结果
                var sheet = pck.Workbook.Worksheets.First();
                //确保所有的转换均已完成
                sheet.Cells[sheet.Dimension.Address].Any(p => p.Text.Contains("{{")).ShouldBeFalse();
            }
        }

        [Fact(DisplayName = "Excel模板导出Bytes测试（issues#34_2）")]
        public async Task ExportBytesByTemplate_Test1()
        {
            //模板路径
            var tplPath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "ExportTemplates",
                "template.xlsx");
            //创建Excel导出对象
            IExportFileByTemplate exporter = new ExcelExporter();
            //导出路径
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(ExportBytesByTemplate_Test1) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var airCompressors = new List<AirCompressor>
            {
                new AirCompressor()
                {
                    Name = "1#",
                    Manufactor = "111",
                    ExhaustPressure = "0",
                    ExhaustTemperature = "66.7-95",
                    RunningTime = "35251",
                    WarningError = "正常",
                    Status = "开机"
                },
                new AirCompressor()
                {
                    Name = "2#",
                    Manufactor = "222",
                    ExhaustPressure = "1",
                    ExhaustTemperature = "90.7-95",
                    RunningTime = "2222",
                    WarningError = "正常",
                    Status = "开机"
                }
            };

            var afterProcessings = new List<AfterProcessing>
            {
                new AfterProcessing()
                {
                    Name = "1#abababa",
                    Manufactor = "杭州立山",
                    RunningTime = "NaN",
                    WarningError = "故障",
                    Status = "停机"
                }
            };

            var suggests = new List<Suggest>
            {
                new Suggest()
                {
                    Number = 1,
                    Description = "故障停机",
                    SuggestMessage = "顾问团队远程协助"
                }
            };

            //根据模板导出
            var result = await exporter.ExportBytesByTemplate(
                new ReportInformation()
                {
                    Contacts = "11112",
                    ContactsNumber = "13642666666",
                    CustomerName = "ababace",
                    Date = DateTime.Now.ToString("yyyy年MM月dd日"),
                    SystemExhaustPressure = "0.54-0.62",
                    SystemDewPressure = "-0.63--77.5",
                    SystemDayFlow = "201864",
                    AirCompressors = airCompressors,
                    AfterProcessings = afterProcessings,
                    Suggests = suggests,
                    SystemPressureHisotries = new List<SystemPressureHisotry>()
                },
                tplPath);
            result.ShouldNotBeNull();
            using (var file = File.OpenWrite(filePath))
            {
                file.Write(result, 0, result.Length);
            }

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //检查转换结果
                var sheet = pck.Workbook.Worksheets.First();
                //确保所有的转换均已完成
                sheet.Cells[sheet.Dimension.Address].Any(p => p.Text.Contains("{{")).ShouldBeFalse();
            }
        }

        #endregion 模板导出

    }
}
