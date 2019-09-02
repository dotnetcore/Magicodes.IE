using System.Collections.Generic;

namespace Magicodes.ExporterAndImporter.Core.Models
{
    /// <summary>
    /// 验证结果
    /// </summary>
    public class ValidationResultModel
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="ValidationResultModel"/> class.
        /// </summary>
        public ValidationResultModel()
        {
            Errors = new Dictionary<string, string>();
            FieldErrors = new Dictionary<string, string>();
        }

        /// <summary>
        ///     Gets or sets the index.
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        ///     Gets or sets the errors.
        /// </summary>
        public IDictionary<string, string> Errors { get; set; }

        /// <summary>
        ///     Gets or sets the field errors.
        /// </summary>
        public IDictionary<string, string> FieldErrors { get; set; }
    }
}
