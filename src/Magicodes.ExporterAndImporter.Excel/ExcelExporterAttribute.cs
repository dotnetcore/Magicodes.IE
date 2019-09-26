// ======================================================================
// 
//           Copyright (C) 2019-2030 湖南心莱信息科技有限公司
//           All rights reserved
// 
//           filename : ExcelExporterAttribute.cs
//           description :
// 
//           created by 雪雁 at  2019-09-11 13:51
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using Magicodes.ExporterAndImporter.Core;

namespace Magicodes.ExporterAndImporter.Excel
{
    /// <summary>
    ///     Excel导出设计
    /// </summary>
    public class ExcelExporterAttribute : ExporterAttribute
    {
        /// <summary>
        ///     表格样式风格
        /// </summary>
        public string TableStyle { get; set; } = "Medium10";

        /// <summary>
        ///     自适应所有列
        /// </summary>
        public bool AutoFitAllColumn { get; set; }

        /// <summary>
        ///     作者
        /// </summary>
        public string Author { get; set; }
    }
}