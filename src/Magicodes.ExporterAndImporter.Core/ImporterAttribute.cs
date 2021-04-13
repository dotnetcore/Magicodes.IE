// ======================================================================
//
//           filename : ImporterAttribute.cs
//           description :
//
//           created by 雪雁 at  2019-09-11 13:51
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
//
// ======================================================================

using System;

namespace Magicodes.ExporterAndImporter.Core
{
    /// <summary>
    /// 导入
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class ImporterAttribute : Attribute
    {
        /// <summary>
        ///     表头位置
        /// </summary>
        public int HeaderRowIndex { get; set; } = 1;

        /// <summary>
        /// 最大允许导入的函数
        /// </summary>
        public int MaxCount = 0;

        /// <summary>
        /// 导入结果筛选器
        /// 必须实现【IExporterHeaderFilter】
        /// </summary>
        public Type ImportResultFilter { get; set; }

        /// <summary>
        /// 导入列头筛选器
        /// 必须实现【IImportHeaderFilter】
        /// </summary>
        public Type ImportHeaderFilter { get; set; }

        /// <summary>
        /// 是否禁用所有筛选器
        /// </summary>
        public bool IsDisableAllFilter { get; set; }

        /// <summary>
        /// 是否忽略列的大小写
        /// </summary>
        public bool IsIgnoreColumnCase { get; set; }
    }
}