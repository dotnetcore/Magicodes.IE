// ======================================================================
// 
//           Copyright (C) 2019-2030 湖南心莱信息科技有限公司
//           All rights reserved
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
using Magicodes.ExporterAndImporter.Tests.Models;
using Shouldly;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Xunit;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class ExcelImporter_Tests
    {
        public IImporter Importer = new ExcelImporter();

        [Fact(DisplayName = "生成模板")]
        public async Task GenerateTemplate_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(GenerateTemplate_Test) + ".xlsx");
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            var result = await Importer.GenerateTemplate<ImportProductDto>(filePath);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();

            //TODO:列头隐藏测试
        }

        [Fact(DisplayName = "生成模板字节")]
        public async Task GenerateTemplateBytes_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), nameof(GenerateTemplateBytes_Test) +".xlsx");
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }

            var result = await Importer.GenerateTemplateBytes<ImportProductDto>();
            result.ShouldNotBeNull();
            result.Length.ShouldBeGreaterThan(0);
            File.WriteAllBytes(filePath, result);
            File.Exists(filePath).ShouldBeTrue();
        }

        [Fact(DisplayName = "导入")]
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
                if (item.Name.Contains("空格测试"))
                {
                    item.Name.ShouldBe(item.Name.Trim());
                }

                if (item.Code.Contains("不去除空格测试"))
                {
                    item.Code.ShouldContain(" ");
                }
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

        [Fact(DisplayName = "必填项检测")]
        public async Task IsRequired_Test()
        {
            var pros = typeof(ImportProductDto).GetProperties();
            foreach (var item in pros)
            {
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
        }

        [Fact(DisplayName = "题库导入测试")]
        public async Task QuestionBankImporter_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "题库导入模板.xlsx");
            var import = await Importer.Import<ImportQuestionBankDto>(filePath);
            import.ShouldNotBeNull();

            import.HasError.ShouldBeFalse();
            import.Data.ShouldNotBeNull();
            import.Data.Count.ShouldBe(404);

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
                .ShouldBe(2);

            result.RowErrors.ShouldContain(p => p.RowIndex == 4 && p.FieldErrors.ContainsKey("身份证"));
            result.RowErrors.ShouldContain(p => p.RowIndex == 5 && p.FieldErrors.ContainsKey("身份证"));

            #region 重复错误

            var errorRows = "5,6".Split(',').ToList();
            result.RowErrors.ShouldContain(p =>
                errorRows.Contains(p.RowIndex.ToString()) && p.FieldErrors.ContainsKey("产品代码") &&
                p.FieldErrors.Values.Contains("存在数据重复，请检查！所在行：5，6。"));

            errorRows = "8,9,11,13".Split(',').ToList();
            result.RowErrors.ShouldContain(p =>
                errorRows.Contains(p.RowIndex.ToString()) && p.FieldErrors.ContainsKey("产品代码") &&
                p.FieldErrors.Values.Contains("存在数据重复，请检查！所在行：8，9，11，13。"));

            errorRows = "4，6，8，10，11，13".Split('，').ToList();
            result.RowErrors.ShouldContain(p =>
                errorRows.Contains(p.RowIndex.ToString()) && p.FieldErrors.ContainsKey("产品型号") &&
                p.FieldErrors.Values.Contains("存在数据重复，请检查！所在行：4，6，8，10，11，13。"));
            #endregion

            result.RowErrors.Count.ShouldBeGreaterThan(0);
        }

        [Fact(DisplayName = "模板错误检测")]
        public async Task TplError_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Errors", "模板字段错误.xlsx");
            var result = await Importer.Import<ImportProductDto>(filePath);
            result.ShouldNotBeNull();
            result.HasError.ShouldBeTrue();
            result.TemplateErrors.Count.ShouldBeGreaterThan(0);
            result.TemplateErrors.Count(p => p.ErrorLevel == ErrorLevels.Error).ShouldBe(1);
            result.TemplateErrors.Count(p => p.ErrorLevel == ErrorLevels.Warning).ShouldBe(1);
        }

        [Fact(DisplayName = "截断数据测试")]
        public async Task ImporterDataEnd_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "TestFiles", "Import", "截断数据测试.xlsx");
            var import = await Importer.Import<ImportProductDto>(filePath);
            import.ShouldNotBeNull();
            import.Data.ShouldNotBeNull();
            import.Data.Count.ShouldBe(6);
        }
    }
}