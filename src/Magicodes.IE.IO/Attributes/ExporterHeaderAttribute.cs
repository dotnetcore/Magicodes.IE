using System;

namespace Magicodes.IE.IO
{
    /// <summary>
    /// Controls the header text, number format, and order of an exported column.
    /// </summary>
    /// <remarks>Apply to a property of the exported model. When <see cref="Name"/> is not set, the property name is used as the header.</remarks>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public sealed class ExporterHeaderAttribute : Attribute
    {
        /// <summary>
        /// Gets or sets the column header text shown when exporting. When omitted, the property name is used.
        /// </summary>
        public string? Name { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this column is excluded from the export.
        /// </summary>
        public bool IsIgnore { get; set; }

        /// <summary>
        /// Gets or sets the Excel number format string, for example <c>"0.00"</c> or <c>"yyyy-MM-dd"</c>.
        /// </summary>
        public string? Format { get; set; }

        /// <summary>
        /// Gets or sets the column width. Values less than or equal to 0 leave the width unset.
        /// </summary>
        public double Width { get; set; }

        /// <summary>
        /// Gets or sets the zero-based column order. The default is -1, which preserves the property declaration order.
        /// </summary>
        public int Index { get; set; } = -1;
    }
}
