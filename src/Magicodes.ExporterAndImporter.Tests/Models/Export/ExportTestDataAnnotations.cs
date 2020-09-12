using System.ComponentModel.DataAnnotations;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using System;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    [ExcelExporter(Name = "数据注解测试")]
    public class ExportTestDataAnnotations
    {
        /// <summary>
        /// </summary>
        [Display(Name = "列1")]
        [ExporterHeader("Custom列1")]
        public string Name1 { get; set; }
        [Display(Name = "列2")]
        public string Name2 { get; set; }
        /// <summary>
        /// 时间测试
        /// </summary>
        [DisplayFormat(DataFormatString = "yyyy-MM-dd")]
        public DateTime Time1 { get; set; }
        /// <summary>
        /// 时间测试
        /// </summary>
        [ExporterHeader(Format = "yyyy-MM-dd")]
        public DateTime Time2 { get; set; }
        /// <summary>
        ///     忽略
        /// </summary>
        [IEIgnoreAttribute]
        public string Ignore { get; set; }

        [ValueMapping("A Test", "1")]
        [ValueMapping("B Test", "2")]
        public MyEmum Testa { get; set; }
    }

    public enum MyEmum
    {
        A,
        B
    }
}