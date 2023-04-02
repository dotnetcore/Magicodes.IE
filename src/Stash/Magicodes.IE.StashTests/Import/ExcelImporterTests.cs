using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Magicodes.IE.Stash.Import;
using Magicodes.IE.StashTests.Extensions;

namespace Magicodes.IE.StashTests.Import
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
    }
}