using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Tests.Models.Import;
using Newtonsoft.Json;
using OfficeOpenXml;
using Shouldly;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
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
                if (import.RowErrors.Count > 0)
                    _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
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

                if (import.RowErrors.Count > 0)
                    _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));

                import.Data.ShouldNotBeNull();
                if (item.Key == "1班导入数据")
                {
                    import.Data.Count.ShouldBe(16);
                    ImportStudentDto dto = (ImportStudentDto) import.Data.ElementAt(0);
                    dto.Name.ShouldBe("杨圣超");
                }

                if (item.Key == "1")
                {
                    import.HasError.ShouldBeTrue();
                    import.Data.Count.ShouldBe(20);
                    ImportPaymentLogDto dto = (ImportPaymentLogDto) import.Data.ElementAt(0);
                    dto.Name.ShouldBe("刘茵");
                }
            }
        }


        [Fact(DisplayName = "学生基础数据及缴费流水号导入_标注错误")]
        public async Task ClassStudentInfoImporter_SaveLabelingError_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import",
                "学生基础数据及缴费流水号导入_标注错误.xlsx");

            var importDic = await Importer.ImportMultipleSheet<ImportStudentAndPaymentLogDto>(filePath);
            foreach (var item in importDic)
            {
                var import = item.Value;
                import.ShouldNotBeNull();
                if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

                if (import.RowErrors.Count > 0)
                    _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));

                import.Data.ShouldNotBeNull();
                if (item.Key == "1班导入数据")
                {
                    import.Data.Count.ShouldBe(16);
                    ImportStudentDto dto = (ImportStudentDto) import.Data.ElementAt(0);
                    dto.Name.ShouldBe("杨圣超");
                }

                if (item.Key == "1")
                {
                    import.HasError.ShouldBeTrue();
                    import.Data.Count.ShouldBe(20);
                    ImportPaymentLogDto dto = (ImportPaymentLogDto) import.Data.ElementAt(0);
                    dto.Name.ShouldBe("刘茵");
                }
            }

            var ext = Path.GetExtension(filePath);
            var labelingErrorExcelPath = filePath.Replace(ext, "_" + ext);
            if (File.Exists(labelingErrorExcelPath))
            {
                _testOutputHelper.WriteLine($"保存标注错误Excel文件已生成,路径：{labelingErrorExcelPath}");
            }
        }


        [Fact(DisplayName = "多Sheet导入模板生成")]
        public async Task MultipleSheetGenerateTemplate_Test()
        {
            var importer = new ExcelImporter();
            var result = await importer.GenerateTemplateBytes<ImportClassStudentDto>();
            var filePath = GetTestFilePath($"{nameof(MultipleSheetGenerateTemplate_Test)}.xlsx");
            DeleteFile(filePath);
            result.ShouldNotBeNull();
            result.Length.ShouldBeGreaterThan(0);
            result.ToExcelExportFileInfo(filePath);
            File.Exists(filePath).ShouldBeTrue();

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //检查转换结果
                pck.Workbook.Worksheets.Count.ShouldBe(2);
#if NET461
                pck.Workbook.Worksheets[1].Name.ShouldBe("1班导入数据");
                pck.Workbook.Worksheets[2].Name.ShouldBe("2班导入数据");
#else
                pck.Workbook.Worksheets[0].Name.ShouldBe("1班导入数据");
                pck.Workbook.Worksheets[1].Name.ShouldBe("2班导入数据");
#endif
            }
        }


        [Fact(DisplayName = "学生基础数据及缴费流水号导入_通过流导入")]
        public async Task StudentInfoAndPaymentLogImporterByStream_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "学生基础数据及缴费流水号导入.xlsx");
            using (Stream stream = new FileStream(filePath, FileMode.Open))
            {
                var importDic = await Importer.ImportMultipleSheet<ImportStudentAndPaymentLogDto>(stream);
                foreach (var item in importDic)
                {
                    var import = item.Value;
                    import.ShouldNotBeNull();
                    if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());
                    if (import.RowErrors.Count > 0)
                        _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors
                        ));
                    import.Data.ShouldNotBeNull();
                    if (item.Key == "1班导入数据")
                    {
                        import.Data.Count.ShouldBe(16);
                        ImportStudentDto dto = (ImportStudentDto) import.Data.ElementAt(0);
                        dto.Name.ShouldBe("杨圣超");
                    }

                    if (item.Key == "缴费数据")
                    {
                        import.HasError.ShouldBeTrue();
                        import.Data.Count.ShouldBe(20);
                        ImportPaymentLogDto dto = (ImportPaymentLogDto) import.Data.ElementAt(0);
                        dto.Name.ShouldBe("刘茵");
                    }
                }
            }
        }
    }
}