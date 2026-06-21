#if NET6_0_OR_GREATER
using System;
using Magicodes.ExporterAndImporter.Core;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import
{
    /// <summary>
    /// #588: DateOnly 类型导入测试 DTO。
    /// </summary>
    public class DateOnlyImportDto
    {
        [ImporterHeader(Name = "名称")]
        public string Name { get; set; }

        [ImporterHeader(Name = "日期")]
        public DateOnly Date { get; set; }

        [ImporterHeader(Name = "可选日期")]
        public DateOnly? NullableDate { get; set; }
    }
}
#endif