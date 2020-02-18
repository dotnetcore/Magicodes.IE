using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using Newtonsoft.Json;
using Shouldly;
using System.IO;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;

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





    }
}
