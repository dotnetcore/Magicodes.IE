// ======================================================================
// 
//           filename : DataRowErrorInfo.cs
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
    ///     数据行错误信息
    /// </summary>
    public class DataRowErrorInfo
    {
        /// <summary>
        ///     Initializes a new instance of the <see cref="DataRowErrorInfo" /> class.
        /// </summary>
        public DataRowErrorInfo()
        {
            FieldErrors = new Dictionary<string, string>();
        }

        /// <summary>
        ///     序号
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        ///     字段错误信息
        /// </summary>
        public IDictionary<string, string> FieldErrors { get; set; }
    }
}