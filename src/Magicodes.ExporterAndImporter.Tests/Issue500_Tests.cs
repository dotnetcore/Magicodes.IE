using DocumentFormat.OpenXml.Bibliography;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Tests.Models.Import;
using Magicodes.IE.Tests.Models.Import;
using Newtonsoft.Json;
using OfficeOpenXml;
using Shouldly;
using System;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class ExcelImporter_Tests_500 : TestBase
    {
        public ExcelImporter_Tests_500(ITestOutputHelper testOutputHelper)
        {
            _testOutputHelper = testOutputHelper;
        }

        private readonly ITestOutputHelper _testOutputHelper;
        public IExcelImporter Importer = new ExcelImporter();

        [Fact(DisplayName = "带导入说明行的 Excel 自动标注错误后导出文件里表头缺失")]
        public async Task Issue500_Test1()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "issue500带描述头测试.xlsx");

            var errorStream = new MemoryStream();
            var fs = new FileStream(filePath, FileMode.Open);

            var import = await Importer.Import<Issue500>(fs, errorStream);

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));

            using (var pck = new ExcelPackage(errorStream))
            {
                pck.Workbook.Worksheets.Count.ShouldBe(1);
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Cells[1, 1].Text.ShouldBe("序号");
                sheet.Cells[2, 1].Text.ShouldNotBeNullOrWhiteSpace();
            }
        }

        //[Fact(DisplayName = "异常流增加自定义异常后导出 fileByte 为 null")]
        //public async Task Issue998_Test()
        //{
        //    var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "学生基础数据导入带描述头测试下.xlsx");

        //    var errorStream = new MemoryStream();
        //    var fs = new FileStream(filePath, FileMode.Open);

        //    var import = await Importer.Import<ImportWithOnlyError>(fs, errorStream);

        //    if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));

        //    foreach (var item in import.Data.ToList())
        //    {
        //        var errorInfo = new DataRowErrorInfo()
        //        {
        //            //由于 Index 从开始
        //            RowIndex = import.Data.ToList().FindIndex(o => o.Equals(item)) + 3,
        //        };

        //        errorInfo.FieldErrors.Add("序号", "数据库已重复");

        //        import.RowErrors.Add(errorInfo);
        //    }

        //    bool result = Importer.OutputBussinessErrorData<ImportWithOnlyError>(errorStream, import.RowErrors.ToList(), out byte[] msg);

        //    msg.ShouldNotBeNull();
        //}
    }
}
