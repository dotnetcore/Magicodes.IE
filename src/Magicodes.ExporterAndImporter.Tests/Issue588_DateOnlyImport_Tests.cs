#if NET6_0_OR_GREATER
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Tests.Models.Import;
using Newtonsoft.Json;
using OfficeOpenXml;
using Shouldly;
using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;

namespace Magicodes.ExporterAndImporter.Tests
{
    /// <summary>
    /// #588: DateOnly / DateOnly? 导入。
    /// </summary>
    public class Issue588_DateOnlyImport_Tests : TestBase
    {
        private readonly ITestOutputHelper _output;

        public Issue588_DateOnlyImport_Tests(ITestOutputHelper output)
        {
            _output = output;
        }

        #region 测试数据准备

        /// <summary>
        /// 创建包含日期列的 Excel 测试文件。
        /// Row 1: 表头
        /// Row 2: 名称="张三", 日期=2024-06-15, 可选日期=2024-12-01
        /// Row 3: 名称="李四", 日期=2024-01-01, 可选日期=（空）
        /// </summary>
        private string CreateDateOnlyTestExcel()
        {
            var path = GetTestFilePath($"Issue588_DateOnly_{Guid.NewGuid():N}.xlsx");
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = "名称";
                ws.Cells[1, 2].Value = "日期";
                ws.Cells[1, 3].Value = "可选日期";

                ws.Cells[2, 1].Value = "张三";
                ws.Cells[2, 2].Value = new DateTime(2024, 6, 15);
                ws.Cells[2, 3].Value = new DateTime(2024, 12, 1);

                ws.Cells[3, 1].Value = "李四";
                ws.Cells[3, 2].Value = new DateTime(2024, 1, 1);
                // Row 3 Col 3 intentionally left empty

                pck.SaveAs(new FileInfo(path));
            }
            return path;
        }

        #endregion

        #region DateOnly 导入

        [Fact(DisplayName = "#588 DateOnly 导入：日期列正确读取为 DateOnly")]
        public async Task DateOnlyImport_StandardDate_ShouldReadAsDateOnly()
        {
            var filePath = CreateDateOnlyTestExcel();
            try
            {
                IExcelImporter importer = new ExcelImporter();
                var import = await importer.Import<DateOnlyImportDto>(filePath);

                if (import.Exception != null)
                    _output.WriteLine($"Import exception: {import.Exception}");

                if (import.RowErrors.Count > 0)
                    _output.WriteLine($"Row errors: {JsonConvert.SerializeObject(import.RowErrors)}");

                import.ShouldNotBeNull();
                import.HasError.ShouldBeFalse();
                import.Data.Count.ShouldBe(2);

                var data = import.Data.ToList();

                // Row 1
                data[0].Name.ShouldBe("张三");
                data[0].Date.ShouldBe(new DateOnly(2024, 6, 15));
                data[0].NullableDate.ShouldBe(new DateOnly(2024, 12, 1));

                // Row 2 — NullableDate 为空
                data[1].Name.ShouldBe("李四");
                data[1].Date.ShouldBe(new DateOnly(2024, 1, 1));
                data[1].NullableDate.ShouldBeNull();
            }
            finally
            {
                DeleteFile(filePath);
            }
        }

        #endregion
    }
}
#endif