namespace Magicodes.IE.IO
{
    /// <summary>
    /// Print page setup for a worksheet. Lengths are in inches.
    /// </summary>
    public sealed class PageSetup
    {
        /// <summary>
        /// Gets or sets the paper size code (for example, 1 = Letter).
        /// </summary>
        public int? PaperSize { get; init; }

        /// <summary>
        /// Gets or sets the orientation, for example <c>portrait</c> or <c>landscape</c>.
        /// </summary>
        public string? Orientation { get; init; }
        /// <summary>
        /// Gets or sets the scale percentage.
        /// </summary>
        public int? Scale { get; init; }
        /// <summary>
        /// Gets or sets the number of pages to fit the width to.
        /// </summary>
        public int? FitToWidth { get; init; }
        /// <summary>
        /// Gets or sets the number of pages to fit the height to.
        /// </summary>
        public int? FitToHeight { get; init; }

        /// <summary>
        /// Gets or sets a value indicating whether to print in black and white.
        /// </summary>
        public bool? BlackAndWhite { get; init; }

        /// <summary>
        /// Gets or sets the left margin, in inches.
        /// </summary>
        public double MarginLeft { get; init; } = 0.7;
        /// <summary>
        /// Gets or sets the right margin, in inches.
        /// </summary>
        public double MarginRight { get; init; } = 0.7;
        /// <summary>
        /// Gets or sets the top margin, in inches.
        /// </summary>
        public double MarginTop { get; init; } = 0.75;
        /// <summary>
        /// Gets or sets the bottom margin, in inches.
        /// </summary>
        public double MarginBottom { get; init; } = 0.75;
        /// <summary>
        /// Gets or sets the header margin, in inches.
        /// </summary>
        public double MarginHeader { get; init; } = 0.3;
        /// <summary>
        /// Gets or sets the footer margin, in inches.
        /// </summary>
        public double MarginFooter { get; init; } = 0.3;

        /// <summary>
        /// Gets or sets the header text on odd pages.
        /// </summary>
        public string? OddHeader { get; init; }

        /// <summary>
        /// Gets or sets the footer text on odd pages.
        /// </summary>
        public string? OddFooter { get; init; }

        /// <summary>
        /// Gets or sets the header text on even pages.
        /// </summary>
        public string? EvenHeader { get; init; }

        /// <summary>
        /// Gets or sets the footer text on even pages.
        /// </summary>
        public string? EvenFooter { get; init; }
    }
}
