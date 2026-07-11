
using System;

namespace Magicodes.IE.IO
{

    /// <summary>
    /// A workbook- or sheet-scoped named range.
    /// </summary>
    public sealed class NamedRange
    {

        /// <summary>
        /// Gets the range name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the cell range reference, for example <c>Sheet1!$A$1:$B$10</c>.
        /// </summary>
        public string Ref { get; }

        /// <summary>
        /// Gets the range comment.
        /// </summary>
        public string? Comment { get; }

        /// <summary>
        /// Gets the zero-based local sheet index when scoped to a single sheet; <see langword="null"/> for a workbook scope.
        /// </summary>
        public int? LocalSheetId { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="NamedRange"/> class.
        /// </summary>
        public NamedRange(string name, string ref_, string? comment = null, int? localSheetId = null)
        {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            Ref = ref_ ?? throw new ArgumentNullException(nameof(ref_));
            Comment = comment;
            LocalSheetId = localSheetId;
        }
    }
}
