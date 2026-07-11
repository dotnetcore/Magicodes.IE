using System.ComponentModel;

namespace Magicodes.IE.IO
{
    /// <summary>
    /// Resolved metadata for a single exported column, consumed by the lower-level write APIs.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class ColumnMeta
    {
        /// <summary>
        /// Gets the name of the bound property.
        /// </summary>
        public string PropertyName { get; }
        /// <summary>
        /// Gets the text displayed in the header row.
        /// </summary>
        public string DisplayName { get; }
        /// <summary>
        /// Gets the Excel number format string, or <see langword="null"/> if none.
        /// </summary>
        public string? Format { get; }
        /// <summary>
        /// Gets the column width, or <see langword="null"/> to use the default.
        /// </summary>
        public double? Width { get; }
        /// <summary>
        /// Gets a value indicating whether the column is hidden.
        /// </summary>
        public bool Hidden { get; }
        /// <summary>
        /// Gets the resolved style index.
        /// </summary>
        public int StyleId { get; }
        /// <summary>
        /// Gets the zero-based column position.
        /// </summary>
        public int Index { get; }
        /// <summary>
        /// Gets a value indicating whether the text is bold, or <see langword="null"/> for the default.
        /// </summary>
        public bool? Bold { get; }
        /// <summary>
        /// Gets a value indicating whether the text wraps within the cell, or <see langword="null"/> for the default.
        /// </summary>
        public bool? Wrap { get; }
        /// <summary>
        /// Gets the font size in points, or <see langword="null"/> for the default.
        /// </summary>
        public float? FontSize { get; }
        /// <summary>
        /// Gets a value indicating whether content is horizontally centered, or <see langword="null"/> for the default.
        /// </summary>
        public bool? AutoCenter { get; }
        /// <summary>
        /// Gets the background color as an ARGB hex string, or <see langword="null"/> for none.
        /// </summary>
        public string? BackgroundColor { get; }
        /// <summary>
        /// Gets the font color as an ARGB hex string, or <see langword="null"/> for the default.
        /// </summary>
        public string? FontColor { get; }
        /// <summary>
        /// Gets the font name, or <see langword="null"/> for the default.
        /// </summary>
        public string? FontName { get; }
        /// <summary>
        /// Gets the border style, or <see langword="null"/> for none.
        /// </summary>
        public BorderStyle? BorderStyle { get; }
        /// <summary>
        /// Gets the border color as an ARGB hex string, or <see langword="null"/> for none.
        /// </summary>
        public string? BorderColor { get; }
        /// <summary>
        /// Gets a value indicating whether the text is italic, or <see langword="null"/> for the default.
        /// </summary>
        public bool? Italic { get; }
        /// <summary>
        /// Gets a value indicating whether the text is underlined, or <see langword="null"/> for the default.
        /// </summary>
        public bool? Underline { get; }
        /// <summary>
        /// Gets a value indicating whether the text has strikethrough, or <see langword="null"/> for the default.
        /// </summary>
        public bool? StrikeThrough { get; }
        /// <summary>
        /// Gets the vertical alignment, or <see langword="null"/> for the default.
        /// </summary>
        public VerticalAlignment? VerticalAlignment { get; }
        /// <summary>
        /// Gets the row height in points, or <see langword="null"/> for the default.
        /// </summary>
        public double? RowHeight { get; }
        /// <summary>
        /// Gets the formula template, which may contain a <c>{row}</c> placeholder.
        /// </summary>
        public string? Formula { get; }

        /// <summary>
        /// Creates the resolved metadata for a single exported column. The trailing parameters are optional style overrides; omit them to accept the defaults.
        /// </summary>
        public ColumnMeta(string propertyName, string displayName, string? format, double? width, bool hidden, int styleId, int index,
            bool? bold = null, bool? wrap = null, float? fontSize = null, bool? autoCenter = null,
            string? backgroundColor = null, string? fontColor = null, string? fontName = null,
            BorderStyle? borderStyle = null, string? borderColor = null,
            bool? italic = null, bool? underline = null, bool? strikeThrough = null,
            VerticalAlignment? verticalAlignment = null, double? rowHeight = null, string? formula = null)
        {
            PropertyName = propertyName;
            DisplayName = displayName;
            Format = format;
            Width = width;
            Hidden = hidden;
            StyleId = styleId;
            Index = index;
            Bold = bold;
            Wrap = wrap;
            FontSize = fontSize;
            AutoCenter = autoCenter;
            BackgroundColor = backgroundColor;
            FontColor = fontColor;
            FontName = fontName;
            BorderStyle = borderStyle;
            BorderColor = borderColor;
            Italic = italic;
            Underline = underline;
            StrikeThrough = strikeThrough;
            VerticalAlignment = verticalAlignment;
            RowHeight = rowHeight;
            Formula = formula;
        }
    }
}
