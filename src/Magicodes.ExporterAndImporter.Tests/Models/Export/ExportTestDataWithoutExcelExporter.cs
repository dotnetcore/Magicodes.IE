// ======================================================================
// 
//           filename : ExportTestDataWithAttrs.cs
//           description :
// 
//           created by 雪雁 at  2019-11-05 20:02
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using Magicodes.ExporterAndImporter.Core;
using System;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    public class ExportTestDataWithoutExcelExporter
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
    }
}