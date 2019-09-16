using System.Collections.Generic;

namespace Magicodes.ExporterAndImporter.Core.Models
{
    /// <summary>
    /// 导入模型
    /// </summary>
    public class ImportModel<T>
    {
        /// <summary>
        ///     Gets or sets the data.
        /// </summary>
        public IList<T> Data { get; set; }

        /// <summary>
        ///     Gets or sets the validation results.
        /// </summary>
        public IList<ValidationResultModel> ValidationResults { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this instance has valid template.
        /// </summary>
        /// <value>
        /// <c>true</c> if this instance has valid template; otherwise, <c>false</c>.
        /// </value>
        public bool HasValidTemplate { get; set; }
    }
}