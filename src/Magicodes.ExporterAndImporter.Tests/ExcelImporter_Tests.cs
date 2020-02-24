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

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Tests.Models.Export;
using Magicodes.ExporterAndImporter.Tests.Models.Import;
using Newtonsoft.Json;
using OfficeOpenXml;
using Shouldly;
using Xunit;
using Xunit.Abstractions;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class ExcelImporter_Tests : TestBase
    {
        public ExcelImporter_Tests(ITestOutputHelper testOutputHelper)
        {
            _testOutputHelper = testOutputHelper;
        }

        private readonly ITestOutputHelper _testOutputHelper;
        public IImporter Importer = new ExcelImporter();

        /// <summary>
        /// 测试枚举
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "生成学生数据导入模板")]
        public async Task GenerateStudentImportTemplate_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(),
                nameof(GenerateStudentImportTemplate_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var result = await Importer.GenerateTemplate<ImportStudentDto>(filePath);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();

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
                nameof(GenerateStudentImportTemplate_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var result = await Importer.GenerateTemplate<ImportStudentDtoWithSheetDesc>(filePath);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();

            //TODO:读取Excel检查表头和格式
        }

        [Fact(DisplayName = "生成模板")]
        public async Task GenerateTemplate_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(GenerateTemplate_Test) + ".xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var result = await Importer.GenerateTemplate<ImportProductDto>(filePath);
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
            var import = await Importer.Import<ImportProductDto>(filePath);
            import.ShouldNotBeNull();

            import.HasError.ShouldBeFalse();
            import.Data.ShouldNotBeNull();
            import.Data.Count.ShouldBeGreaterThanOrEqualTo(2);
            foreach (var item in import.Data)
            {
                if (item.Name.Contains("空格测试")) item.Name.ShouldBe(item.Name.Trim());

                if (item.Code.Contains("不去除空格测试")) item.Code.ShouldContain(" ");
                //去除中间空格测试
                item.BarCode.ShouldBe("123123");
            }

            //可为空类型测试
            import.Data.ElementAt(4).Weight.HasValue.ShouldBe(true);
            import.Data.ElementAt(5).Weight.HasValue.ShouldBe(false);
            //提取性别公式测试
            import.Data.ElementAt(0).Sex.ShouldBe("女");
            //获取当前日期以及日期类型测试  如果时间不对，请打开对应的Excel即可更新为当前时间，然后再运行此单元测试
            //import.Data[0].FormulaTest.Date.ShouldBe(DateTime.Now.Date);
            //数值测试
            import.Data.ElementAt(0).DeclareValue.ShouldBe(123123);
            import.Data.ElementAt(0).Name.ShouldBe("1212");
            import.Data.ElementAt(0).BarCode.ShouldBe("123123");
            import.Data.ElementAt(1).Name.ShouldBe("12312312");
            import.Data.ElementAt(2).Name.ShouldBe("左侧空格测试");
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

            #endregion

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
                p.FieldErrors.Values.Contains("存在数据重复，请检查！所在行：5，6。"));

            errorRows = new List<int>()
            {
                8,9,11,13
            };
            result.RowErrors.ShouldContain(p =>
                errorRows.Contains(p.RowIndex) && p.FieldErrors.ContainsKey("产品代码") &&
                p.FieldErrors.Values.Contains("存在数据重复，请检查！所在行：8，9，11，13。"));

            errorRows = new List<int>()
            {
                4,6,8,10,11,13
            };
            result.RowErrors.ShouldContain(p =>
                errorRows.Contains(p.RowIndex) && p.FieldErrors.ContainsKey("产品型号") &&
                p.FieldErrors.Values.Contains("存在数据重复，请检查！所在行：4，6，8，10，11，13。"));

            #endregion

            result.RowErrors.Count.ShouldBeGreaterThan(0);

            //一行仅允许存在一条数据
            foreach (var item in result.RowErrors.GroupBy(p => p.RowIndex).Select(p => new { p.Key, Count = p.Count() }))
                item.Count.ShouldBe(1);

        }

        [Fact(DisplayName = "结果筛选器测试")]
        public async Task ImportResultFilter_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Errors", "数据错误.xlsx");
            var labelingFilePath = Path.Combine(Directory.GetCurrentDirectory(), $"{nameof(ImportResultFilter_Test)}.xlsx");
            var result = await Importer.Import<ImportResultFilterDataDto1>(filePath, labelingFilePath);
            File.Exists(labelingFilePath).ShouldBeTrue();
            result.ShouldNotBeNull();
            result.HasError.ShouldBeTrue();
            result.Exception.ShouldBeNull();

            result.TemplateErrors.Count.ShouldBe(0);

            var errorRows = new List<int>()
            {
                5,6
            };
            result.RowErrors.ShouldContain(p =>
                errorRows.Contains(p.RowIndex) && p.FieldErrors.ContainsKey("产品代码") &&
                p.FieldErrors.Values.Contains("Duplicate data exists, please check! Where:5，6。"));

            //TODO:检查标注

        }

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
            importResult.Exception.Message.ShouldContain("最大允许导入条数不能超过");

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
        [Fact(DisplayName = "导入列头筛选器测试带头部描述")]
        public async Task ImportFailureData()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "学生基础数据导入带描述头.xlsx");
            var import = await Importer.Import<ImportStudentDtoWithSheetDesc>(filePath);
            import.ShouldNotBeNull();
            if (import.Exception != null) _testOutputHelper.WriteLine(import.Exception.ToString());

            if (import.RowErrors.Count > 0) _testOutputHelper.WriteLine(JsonConvert.SerializeObject(import.RowErrors));
            import.HasError.ShouldBeFalse();
            import.Data.ShouldNotBeNull();
            import.Data.Count.ShouldBe(16);

            List<DataRowErrorInfo> ErrorList = new List<DataRowErrorInfo>();

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
                ErrorList.Add(errorInfo);
            }

            bool result = Importer.OutputBussinessErrorData<ImportStudentDtoWithSheetDesc>(filePath, ErrorList, out string msg);

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

            List<DataRowErrorInfo> ErrorList = new List<DataRowErrorInfo>();

            //出现五条无法完成业务效验的错误数据
            foreach (var item in import.Data.ToList())
            {

                var errorInfo = new DataRowErrorInfo()
                {
                    //由于 Index 从开始
                    RowIndex = import.Data.ToList().FindIndex(o => o.Equals(item)) + 1,

                };
                errorInfo.FieldErrors.Add("序号", "数据库已重复");
                errorInfo.FieldErrors.Add("学籍号", "无效的学籍号,疑似外来人物");
                ErrorList.Add(errorInfo);
            }
            var result = Importer.OutputBussinessErrorData<ImportStudentDto>(filePath, ErrorList, out string errorDataFilePath);
            result.ShouldBeTrue();

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


            List<DataRowErrorInfo> ErrorList = new List<DataRowErrorInfo>();

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
                ErrorList.Add(errorInfo);
            }
            var result = Importer.OutputBussinessErrorData<ImportGalleryAxisDto>(filePath, ErrorList, out string errorDataFilePath);
            result.ShouldBeTrue();

        }

        /// <summary>
        /// 重复标注测试,,想已有标注的模板再次插入标注会报错
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "重复标注测试")]
        public async Task StudentInfoWithCommentImporter_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "学生基础数据导入带描述头_.xlsx");
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
        /// 标注未移除 
        /// </summary>
        /// <returns></returns>
        [Fact(DisplayName = "标注需要手动移除测试")]
        public async Task ImportCommentDidntRemove_Test()
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

    }
}