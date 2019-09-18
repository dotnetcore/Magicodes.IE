namespace Magicodes.ExporterAndImporter.Core.Models
{
    /// <summary>
    /// 模板错误信息
    /// </summary>
    public class TemplateErrorInfo
    {
        /// <summary>
        /// 错误等级
        /// </summary>
        public ErrorLevels ErrorLevel { get; set; }
        /// <summary>
        /// Excel列名
        /// </summary>
        public string ColumnName { get; set; }

        /// <summary>
        /// 要求的列名
        /// </summary>
        public string RequireColumnName { get; set; }

        /// <summary>
        /// 消息
        /// </summary>
        public string Message { get; set; }
    }
}