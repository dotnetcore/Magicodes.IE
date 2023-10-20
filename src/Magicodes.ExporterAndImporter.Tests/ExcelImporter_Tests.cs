// ======================================================================
//
//           filename : ExcelImporter_Tests.cs
//           description :
//
//           created by 雪雁 at  2019-09-11 13:51
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
//
// ======================================================================

using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Tests.Models.Export;
using Magicodes.ExporterAndImporter.Tests.Models.Import;
using Magicodes.IE.Core;
using Magicodes.ExporterAndImporter.Tests.Models.Import.Issue305;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using Shouldly;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Xunit;
using Xunit.Abstractions;
using Magicodes.IE.Tests.Models.Import;
using System.Globalization;
using SixLabors.ImageSharp;
using System.Runtime.InteropServices;
using Codeuctivity.ImageSharpCompare;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class ExcelImporter_Tests : TestBase
    {
        public ExcelImporter_Tests(ITestOutputHelper testOutputHelper)
        {
            _testOutputHelper = testOutputHelper;
        }

        private readonly ITestOutputHelper _testOutputHelper;
        public IExcelImporter Importer = new ExcelImporter();

        /// <summary>
        /// 生成学生数据导入模板
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "生成学生数据导入模板")]
        public async Task GenerateStudentImportTemplate_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                nameof(GenerateStudentImportTemplate_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var result = await Importer.GenerateTemplate<GenerateStudentImportTemplateDto>(filePath);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();

            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                pck.Workbook.Worksheets.Count.ShouldBe(3);
                var sheet = pck.Workbook.Worksheets.First();
                var dataValidataions = sheet.DataValidations.FirstOrDefault()
                    as OfficeOpenXml.DataValidation.ExcelDataValidationList;
                dataValidataions.Formula.ExcelFormula.ShouldBe("hidden_Gender!$A$1:$A$2");
            }
            //TODO:读取Excel检查表头和格式
        }

        /// <summary>
        /// 生成学生数据导入模板带有数据验证
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "生成学生数据导入模板带有数据验证")]
        public async Task GenerateStudentImportSheetDataValidationTemplate_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                nameof(GenerateStudentImportSheetDataValidationTemplate_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var result = await Importer.GenerateTemplate<GenerateStudentImportSheetDataValidationDto>(filePath);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Cells["A2"].Style.Numberformat.Format.ShouldBe("@");
            }
            //TODO:读取Excel检查表头和格式
        }

        /// <summary>
        /// 测试生成导入描述头
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "生成学生数据导入模板加描述")]
        public async Task GenerateStudentImportSheetDescriptionTemplate_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                nameof(GenerateStudentImportSheetDescriptionTemplate_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var result = await Importer.GenerateTemplate<ImportStudentDtoWithSheetDesc>(filePath);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                pck.Workbook.Worksheets.Count.ShouldBe(3);
                var sheet = pck.Workbook.Worksheets.First();
                var attr = typeof(ImportStudentDtoWithSheetDesc).GetAttribute<ExcelImporterAttribute>();
                var text = sheet.Cells["A1"].Text.Replace("\n", string.Empty).Replace("\r", string.Empty);
                text.ShouldBe(attr.ImportDescription.Replace("\n", string.Empty).Replace("\r", string.Empty));
            }
        }

        [Fact(DisplayName = "生成模板")]
        public async Task GenerateTemplate_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(GenerateTemplate_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var result = await Importer.GenerateTemplate<ImportProductDto>(filePath);
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Column(15).Style.Numberformat.Format.ShouldBe("yyyy-MM-dd");
            }
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
            //TODO:读取Excel检查表头和格式
        }

        [Fact(DisplayName = "生成模板字节")]
        public async Task GenerateTemplateBytes_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(GenerateTemplateBytes_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var result = await Importer.GenerateTemplateBytes<ImportProductDto>();
            result.ShouldNotBeNull();
            result.Length.ShouldBeGreaterThan(0);
            File.WriteAllBytes(filePath, result);
            File.Exists(filePath).ShouldBeTrue();
        }

        /// <summary>
        /// https://github.com/dotnetcore/Magicodes.IE/issues/293
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "#293")]
        public async Task GenerateTemplateBytesForIssue293_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(GenerateTemplateBytesForIssue293_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var temp = await Importer.GenerateTemplateBytes<ImportProductDto>();
            var mockeSelctItems = new List<string> { "测试一", "测试二", "测试三" };
            // TODO: 添加真实的下拉数据，添加后， 模板会出问题
            using (var stream = new MemoryStream(temp))
            {
                using (var package = new ExcelPackage(stream))
                {
                    var sheet = package.Workbook.Worksheets.FirstOrDefault();
                    var listSheet = package.Workbook.Worksheets.Add("业务分类");
                    listSheet.Hidden = eWorkSheetHidden.Hidden;
                    int rowIndex = 1;
                    foreach (var item in mockeSelctItems)
                    {
                        listSheet.Cells[rowIndex, 1].Value = item;
                        rowIndex++;
                    }
                    var validation = sheet.DataValidations.AddListValidation("A2:A10000");
                    validation.ShowErrorMessage = false;
                    validation.ErrorStyle = ExcelDataValidationWarningStyle.warning;
                    validation.ErrorTitle = "输入得值无效";
                    validation.Error = "请选择所属分类";
                    //validation.Formula.ExcelFormula = $"=OFFSET(业务分类!$A$1,MATCH(""&A2&"",业务分类!$A:$A,0)-1,,COUNTIF(业务分类!$A:$A,""&A2&""),)";
                    validation.Formula.ExcelFormula = $"业务分类!$A$1:$A${rowIndex - 1}";
                    File.WriteAllBytes(filePath, package.GetAsByteArray());
                    File.Exists(filePath).ShouldBeTrue();
                }
            }
        }

        /// <summary>
        /// 测试：
        /// 表头行位置设置
        /// 导入逻辑测试
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "产品信息导入")]
        public async Task ImporterWithStream_Test()
        {
            //第一列乱序
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "产品导入模板.xlsx");
            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                var result = await Importer.Import<ImportProductDto>(stream);
                result.ShouldNotBeNull();

                result.HasError.ShouldBeTrue();
                result.RowErrors.Count.ShouldBe(1);
                result.Data.ShouldNotBeNull();
                result.Data.Count.ShouldBeGreaterThanOrEqualTo(2);
                foreach (var item in result.Data)
                {
                    if (item.Name != null && item.Name.Contains("空格测试")) item.Name.ShouldBe(item.Name.Trim());

                    if (item.Code.Contains("不去除空格测试")) item.Code.ShouldContain(" ");
                    //去除中间空格测试
                    item.BarCode.ShouldBe("123123");
                }

                //可为空类型测试
                result.Data.ElementAt(4).Weight.HasValue.ShouldBe(true);
                result.Data.ElementAt(5).Weight.HasValue.ShouldBe(false);
                //提取性别公式测试
                result.Data.ElementAt(0).Sex.ShouldBe("女");
                //获取当前日期以及日期类型测试  如果时间不对，请打开对应的Excel即可更新为当前时间，然后再运行此单元测试
                //import.Data[0].FormulaTest.Date.ShouldBe(DateTime.Now.Date);
                //数值测试
                result.Data.ElementAt(0).DeclareValue.ShouldBe(123123);
                result.Data.ElementAt(0).Name.ShouldBe("1212");
                result.Data.ElementAt(0).BarCode.ShouldBe("123123");
                result.Data.ElementAt(0).ProductIdTest1.ShouldBe(Guid.Parse("C2EE3694-959A-4A87-BC8C-4003F6576352"));
                result.Data.ElementAt(0).ProductIdTest2.ShouldBe(Guid.Parse("C2EE3694-959A-4A87-BC8C-4003F6576357"));
                result.Data.ElementAt(1).Name.ShouldBe(null);
                result.Data.ElementAt(2).Name.ShouldBe("左侧空格测试");

                result.ImporterHeaderInfos.ShouldNotBeNull();
                result.ImporterHeaderInfos.Count.ShouldBe(17);
            }
        }

        /// <summary>
        /// 测试：
        /// 表头行位置设置
        /// 导入逻辑测试
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "产品信息导入")]
        public async Task Importer_Test()
        {
            //第一列乱序
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "产品导入模板.xlsx");

            var result = await Importer.Import<ImportProductDto>(filePath);
            result.ShouldNotBeNull();

            result.HasError.ShouldBeTrue();
            result.RowErrors.Count.ShouldBe(1);
            result.Data.ShouldNotBeNull();
            result.Data.Count.ShouldBeGreaterThanOrEqualTo(2);
            foreach (var item in result.Data)
            {
                if (item.Name != null && item.Name.Contains("空格测试")) item.Name.ShouldBe(item.Name.Trim());

                if (item.Code.Contains("不去除空格测试")) item.Code.ShouldContain(" ");
                //去除中间空格测试
                item.BarCode.ShouldBe("123123");
            }

            //可为空类型测试
            result.Data.ElementAt(4).Weight.HasValue.ShouldBe(true);
            result.Data.ElementAt(5).Weight.HasValue.ShouldBe(false);
            //提取性别公式测试
            result.Data.ElementAt(0).Sex.ShouldBe("女");
            //获取当前日期以及日期类型测试  如果时间不对，请打开对应的Excel即可更新为当前时间，然后再运行此单元测试
            //import.Data[0].FormulaTest.Date.ShouldBe(DateTime.Now.Date);
            //数值测试
            result.Data.ElementAt(0).DeclareValue.ShouldBe(123123);
            result.Data.ElementAt(0).Name.ShouldBe("1212");
            result.Data.ElementAt(0).BarCode.ShouldBe("123123");
            result.Data.ElementAt(0).ProductIdTest1.ShouldBe(Guid.Parse("C2EE3694-959A-4A87-BC8C-4003F6576352"));
            result.Data.ElementAt(0).ProductIdTest2.ShouldBe(Guid.Parse("C2EE3694-959A-4A87-BC8C-4003F6576357"));
            result.Data.ElementAt(1).Name.ShouldBe(null);
            result.Data.ElementAt(1).Type.ShouldBe(ImporterProductType.Two);
            result.Data.ElementAt(2).Name.ShouldBe("左侧空格测试");
            result.Data.ElementAt(2).Type.ShouldBe(null);

            result.ImporterHeaderInfos.ShouldNotBeNull();
            result.ImporterHeaderInfos.Count.ShouldBe(17);
        }

        [Fact(DisplayName = "截断数据测试")]
        public async Task ImporterDataEnd_Test()
        {
            //中间空行
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "截断数据测试.xlsx");
            var import = await Importer.Import<ImportProductDto2>(filePath);
            import.ShouldNotBeNull();
            import.Data.ShouldNotBeNull();
            import.Data.Count.ShouldBe(6);
        }

        [Fact(DisplayName = "缴费流水导入测试")]
        public async Task ImportPaymentLogs_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "缴费流水导入模板.xlsx");
            var import = await Importer.Import<ImportPaymentLogDto>(filePath);
            import.ShouldNotBeNull();
            import.HasError.ShouldBeTrue();
            import.Exception.ShouldBeNull();
            import.Data.Count.ShouldBe(20);
        }


        [Fact(DisplayName = "导入结果回调函数测试")]
        public async Task ImportResultCallBack_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "缴费流水导入模板.xlsx");
            var import = await Importer.Import<ImportPaymentLogDto>(filePath, (importResult) =>
            {
                int rowNum = 2;//首行数据对应Excel中的行号
                foreach (var importPaymentLogDto in importResult.Data)
                {
                    if (importPaymentLogDto.Amount > 5000)
                    {
                        var dataRowError = new DataRowErrorInfo();
                        dataRowError.RowIndex = rowNum;
                        dataRowError.FieldErrors.Add("金额", "金额不能大于5000");
                        importResult.RowErrors.Add(dataRowError);
                    }
                    rowNum++;
                }
                return importResult;
            });
            import.ShouldNotBeNull();
            import.HasError.ShouldBeTrue();
            import.RowErrors.ShouldContain(p => p.RowIndex == 3 && p.FieldErrors.Values.Contains("金额不能大于5000"));
            import.Exception.ShouldBeNull();
            import.Data.Count.ShouldBe(20);
        }


        [Fact(DisplayName = "必填项检测")]
        public void IsRequired_Test()
        {
            var pros = typeof(ImportProductDto).GetProperties();
            foreach (var item in pros)
                switch (item.Name)
                {
                    //DateTime
                    case "FormulaTest":
                    //int
                    case "DeclareValue":
                    //Required
                    case "Name":
                        item.IsRequired().ShouldBe(true);
                        break;
                    //可为空类型
                    case "Weight":
                    //string
                    case "IdNo":
                        item.IsRequired().ShouldBe(false);
                        break;
                }
        }

        [Fact(DisplayName = "题库导入测试")]
        public async Task QuestionBankImporter_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "题库导入模板.xlsx");
            var import = await Importer.Import<ImportQuestionBankDto>(filePath);
            import.ShouldNotBeNull();
            _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
            import.HasError.ShouldBeFalse();
            import.Data.ShouldNotBeNull();
            import.Data.Count.ShouldBe(404);

            #region 检查Bool值映射

            //是
            import.Data.ElementAt(0).IsDisorderly.ShouldBeTrue();
            //否
            import.Data.ElementAt(1).IsDisorderly.ShouldBeFalse();
            //对
            import.Data.ElementAt(2).IsDisorderly.ShouldBeTrue();
            //错
            import.Data.ElementAt(3).IsDisorderly.ShouldBeFalse();

            #endregion 检查Bool值映射

            import.RowErrors.Count.ShouldBe(0);
            import.TemplateErrors.Count.ShouldBe(0);
        }

        [Fact(DisplayName = "数据错误检测")]
        public async Task RowDataError_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Errors", "数据错误.xlsx");
            var result = await Importer.Import<ImportRowDataErrorDto>(filePath);
            result.ShouldNotBeNull();
            result.HasError.ShouldBeTrue();

            result.TemplateErrors.Count.ShouldBe(0);

            result.RowErrors.ShouldContain(p => p.RowIndex == 2 && p.FieldErrors.ContainsKey("产品名称"));
            result.RowErrors.ShouldContain(p => p.RowIndex == 3 && p.FieldErrors.ContainsKey("产品名称"));

            result.RowErrors.ShouldContain(p => p.RowIndex == 7 && p.FieldErrors.ContainsKey("产品代码"));

            result.RowErrors.ShouldContain(p => p.RowIndex == 3 && p.FieldErrors.ContainsKey("重量(KG)"));
            result.RowErrors.ShouldContain(p => p.RowIndex == 4 && p.FieldErrors.ContainsKey("公式测试"));
            result.RowErrors.ShouldContain(p => p.RowIndex == 5 && p.FieldErrors.ContainsKey("公式测试"));
            result.RowErrors.ShouldContain(p => p.RowIndex == 6 && p.FieldErrors.ContainsKey("公式测试"));
            result.RowErrors.ShouldContain(p => p.RowIndex == 7 && p.FieldErrors.ContainsKey("公式测试"));

            result.RowErrors.ShouldContain(p => p.RowIndex == 3 && p.FieldErrors.ContainsKey("身份证"));
            result.RowErrors.First(p => p.RowIndex == 3 && p.FieldErrors.ContainsKey("身份证")).FieldErrors.Count
                .ShouldBe(3);

            result.RowErrors.ShouldContain(p => p.RowIndex == 4 && p.FieldErrors.ContainsKey("身份证"));
            result.RowErrors.ShouldContain(p => p.RowIndex == 5 && p.FieldErrors.ContainsKey("身份证"));

            #region 重复错误

            var errorRows = new List<int>()
            {
                5,6
            };
            result.RowErrors.ShouldContain(p =>
                errorRows.Contains(p.RowIndex) && p.FieldErrors.ContainsKey("产品代码") &&
                p.FieldErrors.Values.Contains($"{Resource.ExistDuplicateData}5，6。"));

            errorRows = new List<int>()
            {
                8,9,11,13
            };
            result.RowErrors.ShouldContain(p =>
                errorRows.Contains(p.RowIndex) && p.FieldErrors.ContainsKey("产品代码") &&
                p.FieldErrors.Values.Contains($"{Resource.ExistDuplicateData}8，9，11，13。"));

            errorRows = new List<int>()
            {
                4,6,8,10,11,13
            };
            result.RowErrors.ShouldContain(p =>
                errorRows.Contains(p.RowIndex) && p.FieldErrors.ContainsKey("产品型号") &&
                p.FieldErrors.Values.Contains($"{Resource.ExistDuplicateData}4，6，8，10，11，13。"));

            #endregion 重复错误

            result.RowErrors.Count.ShouldBeGreaterThan(0);

            //一行仅允许存在一条数据
            foreach (var item in result.RowErrors.GroupBy(p => p.RowIndex).Select(p => new { p.Key, Count = p.Count() }))
                item.Count.ShouldBe(1);
        }

        //[Fact(DisplayName = "结果筛选器测试")]
        //public async Task ImportResultFilter_Test()
        //{
        //    Importer = new ExcelImporter();
        //    var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Errors", "数据错误.xlsx");
        //    var labelingFilePath = Path.Combine(Directory.GetCurrentDirectory(), $"{nameof(ImportResultFilter_Test)}.xlsx");
        //    var result = await Importer.Import<ImportResultFilterDataDto1>(filePath, labelingFilePath);
        //    File.Exists(labelingFilePath).ShouldBeTrue();
        //    result.ShouldNotBeNull();
        //    result.HasError.ShouldBeTrue();
        //    result.Exception.ShouldBeNull();
        //    result.ImporterHeaderInfos.ShouldNotBeNull();
        //    result.ImporterHeaderInfos.Count.ShouldBeGreaterThan(0);

        //    result.TemplateErrors.Count.ShouldBe(0);

        //    var errorRows = new List<int>()
        //    {
        //        5,6
        //    };
        //    result.RowErrors.ShouldContain(p =>
        //        errorRows.Contains(p.RowIndex) && p.FieldErrors.ContainsKey("产品代码") &&
        //        p.FieldErrors.Values.Contains("Duplicate data exists, please check! Where:5，6。"));

        //    //TODO:检查标注
        //}

        [Fact(DisplayName = "学生基础数据导入")]
        public async Task StudentInfoImporter_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "学生基础数据导入.xlsx");
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

        /// <summary>
        /// https://github.com/dotnetcore/Magicodes.IE/issues/35
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "单列数据导入测试")]
        public async Task OneColumnImporter_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "单列导入测试.xlsx");
            var import = await Importer.Import<Models.Import.OneColumnImporter_Test.OneColumnImporterDto>(filePath);
            import.ShouldNotBeNull();
            import.HasError.ShouldBeFalse();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
            import.HasError.ShouldBeFalse();
            import.Data.ShouldNotBeNull();
            import.Data.Count.ShouldBe(16);
        }

        [Fact(DisplayName = "模板错误检测")]
        public async Task TplError_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Errors", "模板字段错误.xlsx");
            var result = await Importer.Import<ImportProductDto2>(filePath);
            result.ShouldNotBeNull();
            result.HasError.ShouldBeTrue();
            result.TemplateErrors.Count.ShouldBeGreaterThan(0);
            result.TemplateErrors.Count(p => p.ErrorLevel == ErrorLevels.Error).ShouldBe(1);
            result.TemplateErrors.Count(p => p.ErrorLevel == ErrorLevels.Warning).ShouldBe(1);
        }

        [Fact(DisplayName = "大量数据导出并导入")]
        public async Task LargeDataImport_Test()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(LargeDataImport_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var result = await exporter.Export(filePath, GenFu.GenFu.ListOf<ExportTestData>(50001));
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();

            var importResult = await Importer.Import<ExportTestData>(filePath);
            importResult.HasError.ShouldBeTrue();
            importResult.Exception.ShouldNotBeNull();
            //默认最大5万
            importResult.Exception.Message.ShouldContain(Resource.MaximumNumberImportsCannotExceeded);

            if (File.Exists(filePath)) File.Delete(filePath);
            result = await exporter.Export(filePath, GenFu.GenFu.ListOf<ExportTestData>(50000));
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
            importResult = await Importer.Import<ExportTestData>(filePath);
            importResult.HasError.ShouldBeFalse();
        }

        [Fact(DisplayName = "导入列头筛选器测试")]
        public async Task ImportHeaderFilter_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "导入列头筛选器测试.xlsx");
            var import = await Importer.Import<ImportHeaderFilterDataDto1>(filePath);
            import.ShouldNotBeNull();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
            import.HasError.ShouldBeFalse();
            import.Data.ShouldNotBeNull();
            import.Data.Count.ShouldBe(16);

            //检查值映射
            for (int i = 0; i < import.Data.Count; i++)
            {
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

        [Fact(DisplayName = "学生基础数据导入带头部描述")]
        public async Task StudentInfoWithDescImporter_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "学生基础数据导入带描述头.xlsx");
            var import = await Importer.Import<ImportStudentDtoWithSheetDesc>(filePath);
            import.ShouldNotBeNull();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
            import.HasError.ShouldBeFalse();
            import.Data.ShouldNotBeNull();
            import.Data.Count.ShouldBe(16);

            //检查值映射
            for (int i = 0; i < import.Data.Count; i++)
            {
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

        /// <summary>
        /// 使用错误数据按照导入模板导出
        /// 场景说明 使用导入方法且 导入数据验证无问题后 进行业务判断出现错误,手动将错误的数据标记在原来导入的Excel中
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "自定义校验后写入错误数据验证注释")]
        public async Task ImportFailureData()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "学生基础数据导入带描述头.xlsx");
            var import = await Importer.Import<ImportStudentDtoWithSheetDesc>(filePath);
            import.ShouldNotBeNull();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
            if (import.TemplateErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.TemplateErrors));
            import.HasError.ShouldBeFalse();
            import.Data.ShouldNotBeNull();
            import.Data.Count.ShouldBe(16);

            List<DataRowErrorInfo> errorList = new List<DataRowErrorInfo>();

            //出现五条无法完成业务效验的错误数据
            foreach (var item in import.Data.Skip(5).ToList())
            {
                var errorInfo = new DataRowErrorInfo()
                {
                    //由于 Index 从开始
                    RowIndex = import.Data.ToList().FindIndex(o => o.Equals(item)) + 1,
                };
                errorInfo.FieldErrors.Add("序号", "数据库已重复");
                errorInfo.FieldErrors.Add("学籍号", "无效的学籍号,疑似外来人物");
                errorList.Add(errorInfo);
            }

            bool result = Importer.OutputBussinessErrorData<ImportStudentDtoWithSheetDesc>(filePath, errorList, out string msg);

            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                var resultByte = Importer.OutputBussinessErrorData<ImportStudentDtoWithSheetDesc>(stream, errorList, out byte[] fileByte);
                resultByte.ShouldBeTrue();
                fileByte.ShouldNotBeNull();
            }
            result.ShouldBeTrue();
        }

        /// <summary>
        /// 使用错误数据按照导入模板导出
        /// 场景说明 使用导入方法且 导入数据验证无问题后 进行业务判断出现错误,手动将错误的数据标记在原来导入的Excel中
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "导入列头筛选器测试 不带头部描述")]
        public async Task ImportFailureDataWithoutDesc()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "学生基础数据导入.xlsx");
            var import = await Importer.Import<ImportStudentDto>(filePath);
            import.ShouldNotBeNull();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
            import.HasError.ShouldBeFalse();
            import.Data.ShouldNotBeNull();
            import.Data.Count.ShouldBe(16);

            List<DataRowErrorInfo> errorList = new List<DataRowErrorInfo>();

            //出现五条无法完成业务效验的错误数据
            foreach (var item in import.Data.ToList())
            {
                var errorInfo = new DataRowErrorInfo()
                {
                    //由于 Index 从1开始
                    RowIndex = import.Data.ToList().FindIndex(o => o.Equals(item)) + 2,
                };
                errorInfo.FieldErrors.Add("序号", "数据库已重复");
                errorInfo.FieldErrors.Add("学籍号", "无效的学籍号,疑似外来人物");
                errorList.Add(errorInfo);
            }
            var result = Importer.OutputBussinessErrorData<ImportStudentDto>(filePath, errorList, out string errorDataFilePath);
            result.ShouldBeTrue();
        }

        /// <summary>
        ///     仅导出错误列测试
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "仅导出错误列测试")]
        public async Task ImportOnlyErrorRows()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "过滤学生基础数据导入.xlsx");
            var labelErrorFilePath =
                Path.Combine(Directory.GetCurrentDirectory(), nameof(ImportOnlyErrorRows) + ".xlsx");
            if (File.Exists(labelErrorFilePath)) File.Delete(labelErrorFilePath);
            var import = await Importer.Import<ImportWithOnlyErrorRowsDto>(filePath, labelErrorFilePath);
            import.ShouldNotBeNull();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));

            import.RowErrors.ShouldContain(p => p.RowIndex == 2 && p.FieldErrors.ContainsKey("身份证号"));
            import.RowErrors.ShouldContain(p => p.RowIndex == 3 && p.FieldErrors.ContainsKey("身份证号"));
            import.RowErrors.ShouldContain(p => p.RowIndex == 13 && p.FieldErrors.ContainsKey("身份证号"));

            import.HasError.ShouldBeTrue();
            import.RowErrors.Count.ShouldBe(3);

            var errorRows = import.Data.Where((x, i) => import.RowErrors.Any(err => err.RowIndex == i + 2)).ToList();

            var labelErrorImport = await Importer.Import<ImportWithOnlyErrorRowsDto>(labelErrorFilePath);
            labelErrorImport.ShouldNotBeNull();
            if (labelErrorImport.Exception != null) _testOutputHelper.WriteLine(labelErrorImport.Exception.ToString());

            if (labelErrorImport.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(labelErrorImport.RowErrors));

            labelErrorImport.Data.Count.ShouldBe(3);
            labelErrorImport.RowErrors.Count.ShouldBe(3);
            foreach (var errorRow in errorRows)
            {
                labelErrorImport.Data.ShouldContain(x => x.Name == errorRow.Name && x.IdCard == errorRow.IdCard);
            }
        }

        /// <summary>
        ///     仅导出错误列测试流返回
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "仅导出错误列测试流返回")]
        public async Task ImportOnlyErrorRowsByStream()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "过滤学生基础数据导入.xlsx");
            using var errorStream = new MemoryStream();
            var fs = new FileStream(filePath, FileMode.Open);
            var import = await Importer.Import<ImportWithOnlyErrorRowsDto>(fs, errorStream);
            import.ShouldNotBeNull();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));

            import.RowErrors.ShouldContain(p => p.RowIndex == 2 && p.FieldErrors.ContainsKey("身份证号"));
            import.RowErrors.ShouldContain(p => p.RowIndex == 3 && p.FieldErrors.ContainsKey("身份证号"));
            import.RowErrors.ShouldContain(p => p.RowIndex == 13 && p.FieldErrors.ContainsKey("身份证号"));

            import.HasError.ShouldBeTrue();
            import.RowErrors.Count.ShouldBe(3);

            var errorRows = import.Data.Where((x, i) => import.RowErrors.Any(err => err.RowIndex == i + 2)).ToList();

            errorStream.Seek(0, SeekOrigin.Begin);
            var labelErrorImport = await Importer.Import<ImportWithOnlyErrorRowsDto>(errorStream);
            labelErrorImport.ShouldNotBeNull();
            if (labelErrorImport.Exception != null) _testOutputHelper.WriteLine(labelErrorImport.Exception.ToString());

            if (labelErrorImport.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(labelErrorImport.RowErrors));

            labelErrorImport.Data.Count.ShouldBe(3);
            labelErrorImport.RowErrors.Count.ShouldBe(3);
            foreach (var errorRow in errorRows)
            {
                labelErrorImport.Data.ShouldContain(x => x.Name == errorRow.Name && x.IdCard == errorRow.IdCard);
            }
        }

        /// <summary>
        /// 管轴导入测试 测试能否手动新增错误信息
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "管轴导入测试")]
        public async Task ImportFailureAxisDataWithoutDesc()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "管轴导入数据.xlsx");
            var import = await Importer.Import<ImportGalleryAxisDto>(filePath);
            import.ShouldNotBeNull();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
            import.HasError.ShouldBeFalse();
            import.Data.ShouldNotBeNull();

            List<DataRowErrorInfo> errorList = new List<DataRowErrorInfo>();

            //出现五条无法完成业务效验的错误数据
            foreach (var item in import.Data.ToList())
            {
                var errorInfo = new DataRowErrorInfo()
                {
                    //由于 Index 从开始
                    RowIndex = import.Data.ToList().FindIndex(o => o.Equals(item)) + 1,
                };
                errorInfo.FieldErrors.Add("管轴编号", "数据库已重复");
                errorInfo.FieldErrors.Add("管廊编号", "责任区域不存在");
                errorInfo.FieldErrors.Add("责任区域", "责任区域不存在");
                errorList.Add(errorInfo);
            }
            var result = Importer.OutputBussinessErrorData<ImportGalleryAxisDto>(filePath, errorList, out string errorDataFilePath);
            result.ShouldBeTrue();
        }

        /// <summary>
        ///     导入带有空行的测试
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "导入带有空行的测试")]
        public async Task ImportNullRows_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "管轴导入数据带有空行.xlsx");
            var import = await Importer.Import<ImportGalleryAxisDto>(filePath);
            import.ShouldNotBeNull();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
            import.HasError.ShouldBeTrue();
            import.Data.ShouldNotBeNull();
        }

        /// <summary>
        /// 标注未移除
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "标注需要手动移除测试")]
        public async Task ImportComment_Test()
        {
            //存在四条重复的学籍号码 ,我们手动修改了两条错误数据还剩下两条错误数据
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "学生基础数据导入存在问题.xlsx");
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //此时B2到B5 都存在错误标注
                pck.Workbook.Worksheets.First().Cells["B2"].Comment.ShouldNotBeNull();
                pck.Workbook.Worksheets.First().Cells["B3"].Comment.ShouldNotBeNull();
                pck.Workbook.Worksheets.First().Cells["B4"].Comment.ShouldNotBeNull();
                pck.Workbook.Worksheets.First().Cells["B5"].Comment.ShouldNotBeNull();
            }

            var import = await Importer.Import<ImportStudentDto>(filePath);
            import.ShouldNotBeNull();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
            import.HasError.ShouldBeTrue();
            import.RowErrors.Count.ShouldBe(2);

            var ext = Path.GetExtension(filePath);
            filePath = filePath.Replace(ext, "_" + ext);

            //此处断点可以发现Excel依然存在4个标注
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                //检查忽略列
                pck.Workbook.Worksheets.First().Cells["B2"].Comment.ShouldNotBeNull();
                pck.Workbook.Worksheets.First().Cells["B3"].Comment.ShouldNotBeNull();

                //这个时候 B4 B5 上面的标注应该去掉
                pck.Workbook.Worksheets.First().Cells["B4"].Comment.ShouldBeNull();
                pck.Workbook.Worksheets.First().Cells["B5"].Comment.ShouldBeNull();
            }
        }

        #region 图片测试

        [Fact(DisplayName = "导入图片测试")]
        public async Task ImportPicture_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "图片导入模板.xlsx");
            var import = await Importer.Import<ImportPictureDto>(filePath);

            import.ShouldNotBeNull();
            import.HasError.ShouldBeFalse();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
            foreach (var item in import.Data)
            {
                if (item.Img != null)
                    File.Exists(item.Img).ShouldBeTrue();
                File.Exists(item.Img1).ShouldBeTrue();
            }

            import.Data.ElementAt(3).Img.ShouldBeNullOrEmpty();
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                //添加严格校验，防止图片位置错误等问题
                var png1 = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Images", "1.Jpeg");
                var png2 = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Images", "3.Jpeg");
                var png3 = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Images", "4.Jpeg");
                var png4 = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Images", "2.Jpeg");

                ImageSharpCompare.ImagesAreEqual(import.Data.ElementAt(0).Img1, png1).ShouldBeTrue();
                ImageSharpCompare.ImagesAreEqual(import.Data.ElementAt(0).Img, png2).ShouldBeTrue();

                //ImageSharpCompare.ImagesAreEqual(import.Data.ElementAt(1).Img1, png3).ShouldBeTrue();
                ImageSharpCompare.ImagesAreEqual(import.Data.ElementAt(1).Img, png4).ShouldBeTrue();

                ImageSharpCompare.ImagesAreEqual(import.Data.ElementAt(2).Img1, png1).ShouldBeTrue();
                ImageSharpCompare.ImagesAreEqual(import.Data.ElementAt(2).Img, png1).ShouldBeTrue();
            }
        }

        [Fact(DisplayName = "导入图片测试头部非第一行")]
        public async Task ImportPictureHeaderNotFirstRow_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "图片导入模板非表头非第一行.xlsx");
            var import = await Importer.Import<ImportPictureHeaderNotFirstRowDto>(filePath);

            import.ShouldNotBeNull();
            import.HasError.ShouldBeFalse();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
        }

        [Fact(DisplayName = "导入图片测试头部非第一行并且带有不规整的前几行")]
        public async Task ImportPictureHeaderNotFirstRowAndIrregularRow_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "图片导入模板非表头非第一行并且带有不规整的行.xlsx");
            var import = await Importer.Import<ImportPictureHeaderNotFirstRowDto>(filePath);

            import.ShouldNotBeNull();
            import.HasError.ShouldBeFalse();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
        }

        [Fact(DisplayName = "导入图片测试_base64")]
        public async Task ImportPictureBase64_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "图片导入模板.xlsx");
            var import = await Importer.Import<ImportPictureBase64Dto>(filePath);
            import.ShouldNotBeNull();
            import.HasError.ShouldBeFalse();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
            foreach (var item in import.Data)
            {
                if (item.Img != null)
                    File.Exists(item.Img).ShouldBeTrue();
                item.Img1.ShouldNotBeNull();
            }
        }

        #endregion 图片测试

        [Fact(DisplayName = "导入测试数据注解")]
        public async Task ImportDataAnnotations_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "数据注解测试模板.xlsx");
            var import = await Importer.Import<ImportTestDataAnnotations>(filePath);
            import.ShouldNotBeNull();
            import.HasError.ShouldBeFalse();
            import.Data.ElementAt(0).Name1.ShouldBe(import.Data.ElementAt(0).Name);
        }

        /// <summary>
        /// https://github.com/dotnetcore/Magicodes.IE/issues/182
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "导入数量限制测试")]
        public async Task ImportMaxCountLimit_Test()
        {
            IExporter exporter = new ExcelExporter();
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(ImportMaxCountLimit_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var result = await exporter.Export(filePath, GenFu.GenFu.ListOf<ImportMaxCountLimitDto>(50001));
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();

            var importResult = await Importer.Import<ImportMaxCountLimitDto>(filePath);
            importResult.HasError.ShouldBeFalse();
        }

        [Fact(DisplayName = "ColumnIndex导入测试")]
        public async Task ImportTestColumnIndex_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "ColumnIndex导入测试.xlsx");
            var import = await Importer.Import<ImportTestColumnIndex>(filePath);
            import.HasError.ShouldBeFalse();
            import.TemplateErrors.Count.ShouldBe(0);
            import.ImporterHeaderInfos.Count.ShouldBe(5);
            import.Data.ElementAt(0).Age = 11;
            import.Data.ElementAt(0).ColumnIndexZeroTest = "aaa";
            import.Data.ElementAt(0).ColumnIndexZeroTest2 = "AAA";
        }

        [Fact(DisplayName = "ColumnIndex导入模板生成测试")]
        public async Task ColumnIndexGenerateImportTemplate_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(ColumnIndexGenerateImportTemplate_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var result = await Importer.GenerateTemplate<ImportTestColumnIndex>(filePath);
            using (var pck = new ExcelPackage(new FileInfo(filePath)))
            {
                var sheet = pck.Workbook.Worksheets.First();
                sheet.Cells[1, 2].Value.ShouldBe("Phone");
            }
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
        }

        [Fact(DisplayName = "合并行数据导入")]
        public async Task MergeRowsImportTest()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "合并行.xlsx");
            var import = await Importer.Import<MergeRowsImportDto>(filePath);
            import.ShouldNotBeNull();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
            import.HasError.ShouldBeFalse();
            import.Data.ShouldNotBeNull();
            import.Data.Select(p => p.Sex).Take(8).All(p => p == "男").ShouldBeTrue();
            import.Data.Select(p => p.Sex).Skip(8).All(p => p == "女").ShouldBeTrue();

            import.Data.Select(p => p.Name).Take(3).All(p => p == "张三").ShouldBeTrue();
            import.Data.Select(p => p.Name).Skip(3).Take(4).All(p => p == "李四").ShouldBeTrue();
            import.Data.Select(p => p.Name).Skip(7).Take(6).All(p => p == "王五").ShouldBeTrue();
            import.Data.Count.ShouldBe(13);
        }

        [Fact(DisplayName = "Issue225-枚举测试")]
        public async Task Issue225_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "Issue225.xlsx");
            var import = await Importer.Import<ImportStudentDto>(filePath);
            import.ShouldNotBeNull();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));

            import.RowErrors.Count.ShouldBe(2);
            import.RowErrors[0].FieldErrors.Any(p => p.Value.Contains("不存在模板下拉选项中"));
            import.RowErrors[1].FieldErrors.Any(p => p.Value.Contains("不存在模板下拉选项中"));

            import.HasError.ShouldBeTrue();
            import.Data.ShouldNotBeNull();
            import.Data.Count.ShouldBe(16);

            //检查值映射
            for (int i = 0; i < import.Data.Count; i++)
            {
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

        [Fact(DisplayName = "列头忽略大小写导入测试#243")]
        public async Task IsIgnoreColumnCase_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "Issue243.xlsx");
            var import = await Importer.Import<Issue243>(filePath);
            import.ShouldNotBeNull();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
            import.HasError.ShouldBeFalse();
            import.Data.ShouldNotBeNull();
            import.Data.First().SerialNumber.ShouldNotBe(0);
            import.Data.Count.ShouldBe(16);

            //检查值映射
            for (int i = 0; i < import.Data.Count; i++)
            {
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

        /// <summary>
        /// 单元格合并导入
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "单元格合并导入#305")]
        public async Task CellMergeImport_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "Issue305.xlsx");
            var import = await Importer.Import<ProjectImportDto>(filePath);
            import.ShouldNotBeNull();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
            import.HasError.ShouldBeFalse();
            import.Data.ShouldNotBeNull();

            import.Data.ElementAt(0).CardNumber.ShouldNotBeNullOrEmpty();
            import.Data.ElementAt(0).MonthlyFee.ShouldNotBeNullOrEmpty();
            import.Data.ElementAt(0).Flow.ShouldNotBeNullOrEmpty();

            import.Data.ElementAt(1).CardNumber.ShouldNotBeNullOrEmpty();
            import.Data.ElementAt(1).MonthlyFee.ShouldNotBeNullOrEmpty();
            import.Data.ElementAt(1).Flow.ShouldNotBeNullOrEmpty();
            import.Data.ElementAt(1).PaymentDate.ShouldNotBeNull();

            import.Data.ElementAt(2).CardNumber.ShouldBeNullOrEmpty();
            import.Data.ElementAt(2).MonthlyFee.ShouldBeNullOrEmpty();
            import.Data.ElementAt(2).Flow.ShouldBeNullOrEmpty();

            import.Data.ElementAt(3).CardNumber.ShouldBeNullOrEmpty();
            import.Data.ElementAt(3).MonthlyFee.ShouldBeNullOrEmpty();
            import.Data.ElementAt(3).Flow.ShouldBeNullOrEmpty();

            import.Data.Count.ShouldBe(4);
        }

        [Fact]
        public async Task TimeSpan_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", nameof(TimeSpan_Test) + ".xlsx");
            var import = await Importer.Import<TimeSpanDto>(filePath);
            import.ShouldNotBeNull();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());
            import.HasError.ShouldBeTrue();
            import.RowErrors.Count.ShouldBe(1);
            for (int i = 0; i < import.Data.Count; i++)
            {
                var item = import.Data.ElementAt(i);
                if (i == 0)
                {
                    item.TimeSpan1.ShouldBe(TimeSpan.Parse("1"));
                    item.TimeSpan2.ShouldBe(item.TimeSpan1);
                }
                if (i == 1)
                {
                    item.TimeSpan1.ShouldBe(TimeSpan.Parse("2"));
                    item.TimeSpan2.ShouldBeNull();
                }
                if (i == 3)
                {
                    item.TimeSpan1.ShouldBe(TimeSpan.Parse("6:12"));
                    item.TimeSpan2.ShouldBe(item.TimeSpan1);
                }
            }
        }


        [Fact(DisplayName = "DynamicStringLength")]
        public async Task DynamicStringLength_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "dynamic_stringlength.xlsx");
            var import = await Importer.Import<DynamicStringLengthImportDto>(filePath);

            import.HasError.ShouldBeTrue();
            import.RowErrors.Count.ShouldBe(1);
            import.RowErrors[0].FieldErrors.Count.ShouldBe(1);
            import.RowErrors[0].FieldErrors.ShouldContainKeyAndValue("名称", "名称字数不能超过3");
        }

        [Fact(DisplayName = "RequiredIf")]
        public async Task RequiredIf_Test()
        {

            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "requiredif.xlsx");
            var import = await Importer.Import<RequiredIfAttributeImportDto>(filePath);

            import.HasError.ShouldBeTrue();
            import.RowErrors.Count.ShouldBe(1);
            import.RowErrors[0].FieldErrors.Count.ShouldBe(1);
            import.RowErrors[0].FieldErrors.ShouldContainKeyAndValue("名称", "名称不能为空");

            //其他语言下测试
            System.Threading.Thread.CurrentThread.CurrentCulture = new CultureInfo("en", false);
            Resource.Culture = new CultureInfo("en", false);
            import = await Importer.Import<RequiredIfAttributeImportDto>(filePath);

            import.HasError.ShouldBeTrue();
            import.RowErrors.Count.ShouldBe(1);
            import.RowErrors[0].FieldErrors.Count.ShouldBe(1);
            import.RowErrors[0].FieldErrors.ShouldContainKeyAndValue("名称", "名称不能为空");
        }

        [Fact(DisplayName = "#377-链接导入问题")]
        public async Task Issue377_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "LinkTest.xlsx");
            var import = await Importer.Import<Issue377>(filePath);

            import.HasError.ShouldBeFalse();
            import.Data.Count.ShouldBe(3);
        }
    }
}