using System.Collections.Generic;

namespace Magicodes.ExporterAndImporter.Core.Models
{
    /// <summary>
    /// 数据行错误信息
    /// </summary>
    public class DataRowErrorInfo
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="DataRowErrorInfo"/> class.
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
