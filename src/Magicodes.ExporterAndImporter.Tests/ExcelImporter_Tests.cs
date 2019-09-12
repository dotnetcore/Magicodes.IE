using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Tests.Models;
using Xunit;
using System.IO;
using Shouldly;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class ExcelImporter_Tests
    {
        public IImporter Importer = new ExcelImporter();

        [Fact(DisplayName = "导入")]
        public async Task Importer_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "Models", "产品导入模板.xlsx");
            var import = await Importer.Import<ImportProductDto>(filePath);
            import.ShouldNotBeNull();
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
            //提取性别公式测试
            import.Data[0].Sex.ShouldBe("女");
            //获取当前日期以及日期类型测试  如果时间不对，请打开对应的Excel即可更新为当前时间，然后再运行此单元测试
            import.Data[0].FormulaTest.Date.ShouldBe(DateTime.Now.Date);
            //数值测试
            import.Data[0].DeclareValue.ShouldBe(123123);


        }

        [Fact(DisplayName = "生成模板")]
        public async Task GenerateTemplate_Test()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "testTemplate.xlsx");
            if (File.Exists(filePath)) File.Delete(filePath);

            var result = await Importer.GenerateTemplate<ImportProductDto>(filePath);
            result.ShouldNotBeNull();
            File.Exists(filePath).ShouldBeTrue();
        }
    }
}
