// ======================================================================
// 
//           filename : TemplateErrorInfo.cs
//           description :
// 
//           created by 雪雁 at  2019-09-18 9:43
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

namespace Magicodes.ExporterAndImporter.Core.Models
{
    /// <summary>
    ///     模板错误信息
    /// </summary>
    public class TemplateErrorInfo
    {
        /// <summary>
        ///     错误等级
        /// </summary>
        public ErrorLevels ErrorLevel { get; set; }

        /// <summary>
        ///     Excel列名
        /// </summary>
        public string ColumnName { get; set; }

        /// <summary>
        ///     要求的列名
        /// </summary>
        public string RequireColumnName { get; set; }

        /// <summary>
        ///     消息
        /// </summary>
        public string Message { get; set; }
    }
}