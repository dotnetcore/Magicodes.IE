#if NET6_0_OR_GREATER
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Tests.Models.Import;
using Newtonsoft.Json;
using OfficeOpenXml;
using Shouldly;
using System.IO;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;

namespace Magicodes.ExporterAndImporter.Tests
{
    /// <summary>
    /// #582: 模板 Excel 有空行时，错误提示行号偏移。
    /// </summary>
    public class Issue582_EmptyRowTemplateImport_Tests : TestBase
    {
        private readonly ITestOutputHelper _output;

        public Issue582_EmptyRowTemplateImport_Tests(ITestOutputHelper output)
        {
            _output = output;
        }

        #region 测试数据准备

        /// <summary>
        /// Row 1: 姓名, 年龄, 城市
        /// Row 2: 张三, 25, 北京
        /// Row 3: （空行）
        /// Row 4: （姓名空）, 30, 上海
        /// </summary>
        private string CreateTemplateWithEmptyRow()
        {
            var path = GetTestFilePath($"Issue582_EmptyRow_{System.Guid.NewGuid():N}.xlsx");
            using (var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("Sheet1");
                ws.Cells[1, 1].Value = "姓名";
                ws.Cells[1, 2].Value = "年龄";
                ws.Cells[1, 3].Value = "城市";

                ws.Cells[2, 1].Value = "张三";
                ws.Cells[2, 2].Value = 25;
                ws.Cells[2, 3].Value = "北京";

                // Row 3 intentionally empty

                // Row 4: Name is empty (will fail IsRequired)
                ws.Cells[4, 2].Value = 30;
                ws.Cells[4, 3].Value = "上海";

                pck.SaveAs(new FileInfo(path));
            }
            return path;
        }

        #endregion

        #region 空行模板导入

        [Fact(DisplayName = "#582 模板有空行：错误行号应反映实际行")]
        public async Task EmptyRowTemplateImport_ErrorRowIndex_ShouldMatchActualRow()
        {
            var filePath = CreateTemplateWithEmptyRow();
            try
            {
                IExcelImporter importer = new ExcelImporter();
                var import = await importer.Import<EmptyRowTemplateImportDto>(filePath);

                import.ShouldNotBeNull();
                // Row 3 is empty → skipped. Row 2 (valid) + Row 4 (invalid, missing Name) are both added to Data.
                // The importer adds all non-empty rows to Data regardless of validation errors,
                // with errors tracked separately in RowErrors.
                import.Data.Count.ShouldBe(2);

                // Row 4 has no Name → should report error at Row 4, not Row 3
                import.RowErrors.Count.ShouldBe(1);

                foreach (var rowError in import.RowErrors)
                {
                    _output.WriteLine($"RowError at index {rowError.RowIndex}: {JsonConvert.SerializeObject(rowError.FieldErrors)}");
                }

                // The key assertion: Row 4 has IsRequired violation on "姓名"
                // RowIndex should be 4 (actual row), not 3 (offset due to empty row skip)
                import.RowErrors.ShouldContain(e =>
                    e.FieldErrors.ContainsKey("姓名") &&
                    e.RowIndex == 4);
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