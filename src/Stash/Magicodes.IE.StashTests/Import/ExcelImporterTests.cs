using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Magicodes.IE.Stash.Import;
using Magicodes.IE.StashTests.Extensions;
using Microsoft.VisualStudio.TestPlatform.Utilities;

namespace Magicodes.IE.Stash.Import.Tests
{
    [TestClass()]
    public class ExcelImporterTests
    {
        private readonly ExcelImporter _excelImporter;
        private string resDir = "..\\..\\..\\_res";
        private string defFileName = "正确的定义.xlsx";
        private string dataFileName = "数据源.xlsx";

        public ExcelImporterTests()
        {
            _excelImporter = new ExcelImporter();
        }


        [TestMethod("从文件加载映射定义")]
        public void LoadDefinitionFromExcelFileTest()
        {
            var defPath = Path.Combine(resDir, defFileName);
            var def = _excelImporter.LoadDefinitionFromExcelFile(defPath);
            Console.WriteLine(def.ToJsonString(true));
            Assert.IsNotNull(def);
        }

        [TestMethod("重复变量定义检测")]
        public void BuildTest()
        {
            var defPath = Path.Combine(resDir, "重复变量定义.xlsx");
            var def = _excelImporter.LoadDefinitionFromExcelFile(defPath);
            Assert.ThrowsException<Exception>(() => _excelImporter.Build());
        }

        [TestMethod("编译成功")]
        public void BuildSucceedTest()
        {
            var defPath = Path.Combine(resDir, defFileName);
            var def = _excelImporter.LoadDefinitionFromExcelFile(defPath);
            _excelImporter.Build();
        }


        [TestMethod("正确导入")]
        public void ResolveTest()
        {
            BuildSucceedTest();
            var dataPath = Path.Combine(resDir, dataFileName);
            var output = _excelImporter.Resolve(dataPath);
            Console.WriteLine(output.ToJsonString(true));

            Assert.AreEqual(16, output.Count(), "解析记录数量不对");
        }

        [TestMethod("从文件中指定Sheet加载映射定义")]
        public void LoadDefinitionFromExcelFileTest1()
        {
            var defPath = Path.Combine(resDir, defFileName);
            Assert.ThrowsException<Exception>(() =>
            {
                var def = _excelImporter.LoadDefinitionFromExcelFile(defPath, "我没有");
                //Console.WriteLine(def.ToJsonString(true));
            }, "精确匹配SheetName不通过");


            var def = _excelImporter.LoadDefinitionFromExcelFile(defPath);
            //Console.WriteLine(def.ToJsonString(true));

            Assert.IsTrue(def.Variables.Any(p => p.Name == "Name"), "$definition$ 表中定义的Name变量未发现");
        }
    }
}