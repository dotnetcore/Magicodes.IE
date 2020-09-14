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
    public class CsvImporter_Tests : TestBase
    {
        public CsvImporter_Tests(ITestOutputHelper testOutputHelper)
        {
            _testOutputHelper = testOutputHelper;
        }

        private readonly ITestOutputHelper _testOutputHelper;
        public IImporter Importer = new CsvImporter();


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
                if (i < 1)
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

        [Fact(DisplayName = "学生基础数据导入Stream")]
        public async Task StudentInfoImporterWithStream_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "学生基础数据导入.csv");

            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                var import = await Importer.Import<ImportStudentDto>(stream);
                import.ShouldNotBeNull();
                if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

                if (import.RowErrors.Count > 0)
                    _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
                import.HasError.ShouldBeFalse();
                import.Data.ShouldNotBeNull();
                import.Data.Count.ShouldBe(16);

                //检查值映射
                for (int i = 0; i < import.Data.Count; i++)
                {
                    if (i < 1)
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

        [Fact(DisplayName = "生成模板字节")]
        public async Task GenerateTemplateBytes_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(GenerateTemplateBytes_Test) + ".csv");
            if (File.Exists(filePath)) File.Delete(filePath);

            var result = await Importer.GenerateTemplateBytes<ImportProductDto>();
            result.ShouldNotBeNull();
            result.Length.ShouldBeGreaterThan(0);
            File.WriteAllBytes(filePath, result);
            File.Exists(filePath).ShouldBeTrue();
        }

        [Fact(DisplayName = "生成模板")]
        public async Task GenerateTemplate_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(GenerateTemplate_Test) + ".csv");
            if (File.Exists(filePath)) File.Delete(filePath);

            var result = await Importer.GenerateTemplate<ImportProductDto>(filePath);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();

        }

    }
}
