using System;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using Newtonsoft.Json;
using Shouldly;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Tests.Models.Import;
using Xunit;
using Xunit.Abstractions;
using CsvHelper;
using System.Globalization;
using CsvHelper.Configuration;
using CsvHelper.TypeConversion;
using System.Reflection;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class ExcelCsvImporter_Tests: TestBase
    {
        public ExcelCsvImporter_Tests(ITestOutputHelper testOutputHelper)
        {
            _testOutputHelper = testOutputHelper;
        }

        private readonly ITestOutputHelper _testOutputHelper;
        public IImporter Importer = new ExcelImporter();

       
        [Fact(DisplayName = "单列数据导入测试")]
        public async Task OneColumnImporter_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "单列导入测试.csv");
            var import = await Importer.Import<Models.Import.OneColumnImporter_Test.OneColumnImporterDto>(filePath);
            import.ShouldNotBeNull();
            import.HasError.ShouldBeFalse();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
            import.HasError.ShouldBeFalse();
            import.Data.ShouldNotBeNull();
            import.Data.Count.ShouldBe(10);
        }

        [Fact(DisplayName = "学生基础数据导入")]
        public async Task StudentInfoImporter_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "学生基础数据导入.csv");
            var import = await Importer.Import<ImportStudentDto>(filePath);
            import.ShouldNotBeNull();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
            import.HasError.ShouldBeFalse();
            import.Data.ShouldNotBeNull();
            import.Data.Count.ShouldBe(16);

            //检查值映射
            for (int i = 0; i < import.Data.Count; i++)
            {
                if (i<1)
                {
                    import.Data.ElementAt(i).Status.ShouldBe(StudentStatus.PupilsAway);
                }
                if (i < 5)
                {
                    import.Data.ElementAt(i).Gender.ShouldBe(Genders.Man);
                }
                else
                {
                    import.Data.ElementAt(i).Gender.ShouldBe(Genders.Female);
                }
            }
        }






    }
}
