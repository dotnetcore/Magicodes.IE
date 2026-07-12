using System;
using System.ComponentModel;

namespace Magicodes.IE.IO
{
    /// <summary>
    /// Configuration for a single exported column, created by <see cref="ExportProfile{T}.Column{TProp}"/>. The fluent <c>With*</c> methods return this instance so calls can be chained.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class ColumnConfig
    {
        /// <summary>
        /// Gets or sets the column header text that overrides the inferred name.
        /// </summary>
        public string? Name { get; private set; }
        /// <summary>
        /// Gets or sets the Excel number format string.
        /// </summary>
        public string? Format { get; private set; }
        /// <summary>
        /// Gets or sets the column width.
        /// </summary>
        public double? Width { get; private set; }
        /// <summary>
        /// Gets or sets the zero-based column position.
        /// </summary>
        public int? Index { get; private set; }
        /// <summary>
        /// Gets or sets a value indicating whether the text is bold.
        /// </summary>
        public bool? Bold { get; private set; }
        /// <summary>
        /// Gets or sets a value indicating whether the text wraps within the cell.
        /// </summary>
        public bool? Wrap { get; private set; }
        /// <summary>
        /// Gets or sets a value indicating whether the column is hidden.
        /// </summary>
        public bool? Hidden { get; private set; }
        /// <summary>
        /// Gets or sets the font size in points.
        /// </summary>
        public float? FontSize { get; private set; }
        /// <summary>
        /// Gets or sets a value indicating whether content is horizontally centered.
        /// </summary>
        public bool? AutoCenter { get; private set; }
        /// <summary>
        /// Gets or sets the background color as an ARGB hex string.
        /// </summary>
        public string? BackgroundColor { get; private set; }
        /// <summary>
        /// Gets or sets the font color as an ARGB hex string.
        /// </summary>
        public string? FontColor { get; private set; }
        /// <summary>
        /// Gets or sets the font name.
        /// </summary>
        public string? FontName { get; private set; }
        /// <summary>
        /// Gets or sets the row height in points.
        /// </summary>
        public double? RowHeight { get; private set; }
        /// <summary>
        /// Gets or sets the border style.
        /// </summary>
        public BorderStyle? BorderStyle { get; private set; }
        /// <summary>
        /// Gets or sets the border color as an ARGB hex string.
        /// </summary>
        public string? BorderColor { get; private set; }
        /// <summary>
        /// Gets or sets a value indicating whether the text is italic.
        /// </summary>
        public bool? Italic { get; private set; }
        /// <summary>
        /// Gets or sets a value indicating whether the text is underlined.
        /// </summary>
        public bool? Underline { get; private set; }
        /// <summary>
        /// Gets or sets a value indicating whether the text has strikethrough.
        /// </summary>
        public bool? StrikeThrough { get; private set; }
        /// <summary>
        /// Gets or sets the vertical alignment.
        /// </summary>
        public VerticalAlignment? VerticalAlignment { get; private set; }
        /// <summary>
        /// Gets or sets the formula template, which may contain a <c>{row}</c> placeholder.
        /// </summary>
        public string? Formula { get; private set; }

        /// <summary>
        /// Overrides the column header text. Returns this instance for chaining.
        /// </summary>
        public ColumnConfig WithName(string name) { Name = name; return this; }
        /// <summary>
        /// Sets the number format string. Returns this instance for chaining.
        /// </summary>
        public ColumnConfig WithFormat(string format) { Format = format; return this; }
        /// <summary>
        /// Sets the column width. Returns this instance for chaining.
        /// </summary>
        public ColumnConfig WithWidth(double width) { Width = width; return this; }
        /// <summary>
        /// Sets the zero-based column position. Returns this instance for chaining.
        /// </summary>
        public ColumnConfig WithIndex(int index) { Index = index; return this; }
        /// <summary>
        /// Sets the bold flag. Returns this instance for chaining.
        /// </summary>
        public ColumnConfig WithBold(bool value = true) { Bold = value; return this; }
        /// <summary>
        /// Sets the italic flag. Returns this instance for chaining.
        /// </summary>
        public ColumnConfig WithItalic(bool value = true) { Italic = value; return this; }
        /// <summary>
        /// Sets the underline flag. Returns this instance for chaining.
        /// </summary>
        public ColumnConfig WithUnderline(bool value = true) { Underline = value; return this; }
        /// <summary>
        /// Sets the strikethrough flag. Returns this instance for chaining.
        /// </summary>
        public ColumnConfig WithStrikeThrough(bool value = true) { StrikeThrough = value; return this; }
        /// <summary>
        /// Sets the text-wrap flag. Returns this instance for chaining.
        /// </summary>
        public ColumnConfig WithWrap(bool value = true) { Wrap = value; return this; }
        /// <summary>
        /// Sets the hidden flag. Returns this instance for chaining.
        /// </summary>
        public ColumnConfig WithHidden(bool value = true) { Hidden = value; return this; }
        /// <summary>
        /// Sets the font size. Returns this instance for chaining.
        /// </summary>
        public ColumnConfig WithFontSize(float size) { FontSize = size; return this; }
        /// <summary>
        /// Sets the horizontal-center flag. Returns this instance for chaining.
        /// </summary>
        public ColumnConfig WithAutoCenter(bool value = true) { AutoCenter = value; return this; }
        /// <summary>
        /// Sets the background color as an ARGB hex string. Returns this instance for chaining.
        /// </summary>
        public ColumnConfig WithBackgroundColor(string argb) { BackgroundColor = argb; return this; }
        /// <summary>
        /// Sets the font color as an ARGB hex string. Returns this instance for chaining.
        /// </summary>
        public ColumnConfig WithFontColor(string argb) { FontColor = argb; return this; }
        /// <summary>
        /// Sets the font name. Returns this instance for chaining.
        /// </summary>
        public ColumnConfig WithFontName(string name) { FontName = name; return this; }
        /// <summary>
        /// Sets the data row height. Returns this instance for chaining.
        /// </summary>
        public ColumnConfig WithRowHeight(double height) { RowHeight = height; return this; }
        /// <summary>
        /// Enables or disables a thin border, optionally with a color. Returns this instance for chaining.
        /// </summary>
        public ColumnConfig WithBorder(bool value = true, string? color = null)
        {
            BorderStyle = value ? Magicodes.IE.IO.BorderStyle.Thin : Magicodes.IE.IO.BorderStyle.None;
            BorderColor = color;
            return this;
        }
        /// <summary>
        /// Sets the border style and color. Returns this instance for chaining.
        /// </summary>
        public ColumnConfig WithBorderStyle(BorderStyle style, string? color = null) { BorderStyle = style; BorderColor = color; return this; }
        /// <summary>
        /// Sets the vertical alignment. Returns this instance for chaining.
        /// </summary>
        public ColumnConfig WithVerticalAlignment(VerticalAlignment alignment) { VerticalAlignment = alignment; return this; }
        /// <summary>
        /// Sets the formula template, where the row number is represented by <c>{row}</c>. Returns this instance for chaining.
        /// </summary>
        public ColumnConfig WithFormula(string formulaTemplate) { Formula = formulaTemplate; return this; }
    }
}
