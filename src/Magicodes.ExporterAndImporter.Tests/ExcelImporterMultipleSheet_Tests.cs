using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Tests.Models.Import;
using Newtonsoft.Json;
using Shouldly;
using Xunit;
using Xunit.Abstractions;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class ExcelImporterMultipleSheet_Tests : TestBase
    {

        public ExcelImporterMultipleSheet_Tests(ITestOutputHelper testOutputHelper)
        {
            _testOutputHelper = testOutputHelper;
        }

        private readonly ITestOutputHelper _testOutputHelper;
        public IExcelImporter Importer = new ExcelImporter();


        [Fact(DisplayName = "班级学生基础数据导入")]
        public async Task ClassStudentInfoImporter_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "班级学生基础数据导入.xlsx");
            var importDic = await Importer.ImportSameSheets<ImportClassStudentDto, ImportStudentDto>(filePath);
            foreach (var item in importDic)
            {
                var import = item.Value;
                import.ShouldNotBeNull();
                if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());
                if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
                import.HasError.ShouldBeFalse();
                import.Data.ShouldNotBeNull();
                import.Data.Count.ShouldBe(16);
            }
        }

        [Fact(DisplayName = "学生基础数据及缴费流水号导入")]
        public async Task StudentInfoAndPaymentLogImporter_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "学生基础数据及缴费流水号导入.xlsx");
            var importDic = await Importer.ImportMultipleSheet<ImportStudentAndPaymentLogDto>(filePath);
            foreach (var item in importDic)
            {
                var import = item.Value;
                import.ShouldNotBeNull();
                if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

                if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));

                import.Data.ShouldNotBeNull();
                if (item.Key == "1班导入数据")
                {
                    import.Data.Count.ShouldBe(16);
                    ImportStudentDto dto = (ImportStudentDto)import.Data.ElementAt(0);
                    dto.Name.ShouldBe("杨圣超");
                }
                if (item.Key == "缴费数据")
                {
                    import.HasError.ShouldBeTrue();
                    import.Data.Count.ShouldBe(20);
                    ImportPaymentLogDto dto = (ImportPaymentLogDto)import.Data.ElementAt(0);
                    dto.Name.ShouldBe("刘茵");
                }
            }
        }
    }
}
