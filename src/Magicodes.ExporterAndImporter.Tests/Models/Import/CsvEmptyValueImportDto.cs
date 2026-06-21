using Magicodes.ExporterAndImporter.Core;
using System;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import
{
    /// <summary>
    /// #585: CSV 导入空值处理测试 DTO——空单元格映射为 null 而非空字符串。
    /// </summary>
    public class CsvEmptyValueImportDto
    {
        [ImporterHeader(Name = "姓名")]
        public string Name { get; set; }

        [ImporterHeader(Name = "年龄")]
        public int? Age { get; set; }

        [ImporterHeader(Name = "入职日期")]
        public DateTime? HireDate { get; set; }

        [ImporterHeader(Name = "级别")]
        public string Level { get; set; }
    }
}