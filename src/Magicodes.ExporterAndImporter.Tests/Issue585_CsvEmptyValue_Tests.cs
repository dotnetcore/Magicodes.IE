using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Csv;
using Magicodes.ExporterAndImporter.Tests.Models.Import;
using Newtonsoft.Json;
using Shouldly;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;

namespace Magicodes.ExporterAndImporter.Tests
{
    /// <summary>
    /// #585: CSV 导入空值处理——空单元格映射为 null 而非空字符串。
    /// </summary>
    public class Issue585_CsvEmptyValue_Tests : TestBase
    {
        private readonly ITestOutputHelper _output;

        public Issue585_CsvEmptyValue_Tests(ITestOutputHelper output)
        {
            _output = output;
        }

        #region 测试数据准备

        /// <summary>
        /// CSV 内容：
        ///   姓名,年龄,入职日期,级别
        ///   张三,25,2024-03-15,A
        ///   李四,,,A       ← 年龄和入职日期为空
        ///   王五,30,,
        /// </summary>
        private string CreateEmptyValueCsv()
        {
            var path = GetTestFilePath($"Issue585_EmptyValue_{System.Guid.NewGuid():N}.csv");
            File.WriteAllText(path, "姓名,年龄,入职日期,级别\n张三,25,2024-03-15,A\n李四,,,A\n王五,30,,\n");
            return path;
        }

        #endregion

        #region CSV 空值导入

        [Fact(DisplayName = "#585 CSV 空值导入：空值应为 null 而非空字符串")]
        public async Task CsvImport_EmptyCells_ShouldMapToNull()
        {
            var filePath = CreateEmptyValueCsv();
            try
            {
                IImporter importer = new CsvImporter();
                var import = await importer.Import<CsvEmptyValueImportDto>(filePath);

                if (import.Exception != null)
                    _output.WriteLine($"Exception: {import.Exception}");

                if (import.RowErrors.Count > 0)
                    _output.WriteLine($"RowErrors: {JsonConvert.SerializeObject(import.RowErrors)}");

                import.ShouldNotBeNull();
                import.HasError.ShouldBeFalse();
                import.Data.Count.ShouldBe(3);

                var data = import.Data.ToList();

                // Row 1: all fields populated
                data[0].Name.ShouldBe("张三");
                data[0].Age.ShouldBe(25);
                data[0].Level.ShouldBe("A");

                // Row 2: Age and HireDate empty → should be null, not string.Empty
                data[1].Name.ShouldBe("李四");
                data[1].Age.ShouldBeNull();
                data[1].HireDate.ShouldBeNull();
                data[1].Level.ShouldBe("A");

                // Row 3: Level empty → should be null or empty
                data[2].Name.ShouldBe("王五");
                data[2].Age.ShouldBe(30);
                data[2].Level.ShouldBeNull();
            }
            finally
            {
                DeleteFile(filePath);
            }
        }

        #endregion
    }
}