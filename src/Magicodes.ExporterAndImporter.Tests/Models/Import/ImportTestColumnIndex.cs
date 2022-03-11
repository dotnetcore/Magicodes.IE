using Magicodes.ExporterAndImporter.Core;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import
{
    public class ImportTestColumnIndex
    {
        [ImporterHeader(Name = "姓名")] public string Name { get; set; }

        [ImporterHeader(Name = "列索引测试1")] public string ColumnIndexZeroTest { get; set; }

        [ImporterHeader(ColumnIndex = 2)] public string Phone { get; set; }

        [ImporterHeader(Name = "年龄", ColumnIndex = 3)]
        public int? Age { get; set; }

        public string ColumnIndexZeroTest2 { get; set; }
    }
}