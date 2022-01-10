using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using System;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    [ExcelExporter(Name = "测试", ExcelOutputType = ExcelOutputTypes.None)]
    public class ExportTestDataWithPicture
    {
        [ExporterHeader(DisplayName = "加粗文本", IsBold = true)]
        public string Text { get; set; }

        [ExporterHeader(DisplayName = "普通文本")] public string Text2 { get; set; }
        [ExporterHeader(DisplayName = "忽略", IsIgnore = true)]
        public string Text3 { get; set; }

        [ExportImageField(Width = 20, Height = 120, YOffset = 15)]
        [ExporterHeader(DisplayName = "图1")]
        public string Img1 { get; set; }
        [ExporterHeader(DisplayName = "数值", Format = "#,##0")]
        public decimal Number { get; set; }
        [ExporterHeader(DisplayName = "名称", IsAutoFit = true)]
        public string Name { get; set; }
        /// <summary>
        /// 时间测试
        /// </summary>
        [ExporterHeader(DisplayName = "日期1", Format = "yyyy-MM-dd")]
        public DateTime Time1 { get; set; }

        [ExportImageField(Width = 50, Height = 120, Alt = "404")]
        [ExporterHeader(DisplayName = "图", IsAutoFit = false)]
        public string Img { get; set; }
    }
}
