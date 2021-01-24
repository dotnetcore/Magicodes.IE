using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using OfficeOpenXml.Table;
using System;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    [ExcelExporter(Name = "数据注解测试", TableStyle = TableStyles.Light10)]
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

        [ValueMapping("A Test", "A")]
        [ValueMapping("B Test", "B")]
        public MyEmum MyEmum { get; set; }

        [ValueMapping("是", true)]
        [ValueMapping("否", false)]
        public bool? Bool { get; set; }

        [ValueMapping("是", true)]
        [ValueMapping("否", false)]
        public bool Bool1 { get; set; }
        public bool Bool2 { get; set; }

        /// <summary>
        /// 数值列
        /// </summary>
        [ExporterHeader(DisplayName = "数值", Format = "0")]
        public int? Number { get; set; }
    }

    public enum MyEmum
    {
        A,
        B,
        [Description("C Test")]
        C,
        D
    }
}