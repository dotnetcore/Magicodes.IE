using System;
using System.Collections.Generic;
using System.Text;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    [ExcelExporter(Name = "测试",AutoFitAllColumn = true)]
    public class ExportTestDataWithPicture
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
        [ExporterHeader(DisplayName = "图",IsImg = true,ImgWidth = 50,IsAutoFit = false,ImgHeight = 50)]
        public string Img { get; set; }
    }
}
