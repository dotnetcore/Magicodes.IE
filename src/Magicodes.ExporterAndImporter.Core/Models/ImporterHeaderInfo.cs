// ======================================================================
// 
//           filename : ImporterHeaderInfo.cs
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
using System.Reflection;

namespace Magicodes.ExporterAndImporter.Core.Models
{
    /// <summary>
    ///     导入列头设置
    /// </summary>
    public class ImporterHeaderInfo
    {
        /// <summary>
        ///     是否必填
        /// </summary>
        public bool IsRequired { get; set; }

        /// <summary>
        ///     列名称
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        ///     列属性
        /// </summary>
        public ImporterHeaderAttribute Header { get; set; }
        /// <summary>
        ///     图属性
        /// </summary>
        public ImportImageFieldAttribute ImportImageFieldAttribute { get; set; }

        /// <summary>
        /// </summary>
        public Dictionary<string, dynamic> MappingValues { get; set; } = new Dictionary<string, dynamic>();

        /// <summary>
        ///     是否存在
        /// </summary>
        public bool IsExist { get; set; }

        /// <summary>
        ///     属性信息
        /// </summary>
        public PropertyInfo PropertyInfo { get; set; }
    }
}