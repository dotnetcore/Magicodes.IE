

namespace Magicodes.IE.IO
{

    /// <summary>
    /// Display settings for the worksheet outline (grouping).
    /// </summary>
    public sealed class OutlineSettings
    {

        /// <summary>
        /// Gets or sets a value indicating whether summary rows appear below the detail rows.
        /// </summary>
        public bool SummaryBelow { get; init; } = true;

        /// <summary>
        /// Gets or sets a value indicating whether summary columns appear to the right of the detail columns.
        /// </summary>
        public bool SummaryRight { get; init; } = true;
    }
}
