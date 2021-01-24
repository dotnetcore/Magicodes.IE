using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using OfficeOpenXml.Table;
using System;

namespace MagicodesWebSite
{
    [ExcelExporter(Name = "测试", TableStyle = TableStyles.Light10, AutoFitAllColumn = true)]
    public class ExportTestDataWithAttrs
    {

        [ExporterHeader(DisplayName = "加粗文本", IsBold = true)]
        public string Text { get; set; }
        [ExporterHeader(DisplayName = "普通文本")] public string Text2 { get; set; }
        [ExporterHeader(DisplayName = "忽略", IsIgnore = true)]
        public string Text3 { get; set; }
        [ExporterHeader(DisplayName = "数值", Format = "#,##0")]
        public decimal Number { get; set; }
        [ExporterHeader(DisplayName = "名称", IsAutoFit = true)]
        public string Name { get; set; }

        /// <summary>
        /// 时间测试
        /// </summary>
        [ExporterHeader(DisplayName = "日期1", Format = "yyyy-MM-dd")]
        public DateTime Time1 { get; set; }

        /// <summary>
        /// 时间测试
        /// </summary>
        [ExporterHeader(DisplayName = "日期2", Format = "yyyy-MM-dd HH:mm:ss")]
        public DateTime? Time2 { get; set; }

        public DateTime Time3 { get; set; }

        public DateTime Time4 { get; set; }

        /// <summary>
        /// 长数值测试
        /// </summary>
        [ExporterHeader(DisplayName = "长数值", Format = "#,##0")]
        public long LongNo { get; set; }
    }
}
