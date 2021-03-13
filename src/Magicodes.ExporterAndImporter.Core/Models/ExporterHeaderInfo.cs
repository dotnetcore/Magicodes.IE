// ======================================================================
// 
//           filename : ExporterHeaderInfo.cs
//           description :
// 
//           created by 雪雁 at  2019-09-11 13:51
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using System.Collections.Generic;

namespace Magicodes.ExporterAndImporter.Core.Models
{
    /// <summary>
    /// 导出列头部信息
    /// </summary>
    public class ExporterHeaderInfo
    {
        /// <summary>
        ///     列索引
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        ///     列名称
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        ///     列属性
        /// </summary>
        public ExporterHeaderAttribute ExporterHeaderAttribute { get; set; }

        /// <summary>
        ///     图片属性
        /// </summary>
        public ExportImageFieldAttribute ExportImageFieldAttribute { get; set; }

        /// <summary>
        ///     C#数据类型
        /// </summary>
        public string CsTypeName { get; set; }

        /// <summary>
        ///     最终显示的列名
        /// </summary>
        public string DisplayName { set; get; }

        /// <summary>
        /// </summary>
        public Dictionary<dynamic, string> MappingValues { get; set; } = new Dictionary<dynamic, string>();
    }
}