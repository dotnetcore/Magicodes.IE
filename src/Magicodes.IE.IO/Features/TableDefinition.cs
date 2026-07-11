using System.Collections.Generic;

namespace Magicodes.IE.IO
{
    /// <summary>
    /// A single column within an Excel table.
    /// </summary>
    public sealed class TableColumnDefinition
    {
        /// <summary>
        /// Gets the column name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the totals-row function name, for example <c>sum</c>.
        /// </summary>
        public string? TotalsRowFunction { get; }

        /// <summary>
        /// Creates a column definition.
        /// </summary>
        public TableColumnDefinition(string name, string? totalsRowFunction = null)
        {
            Name = name;
            TotalsRowFunction = totalsRowFunction;
        }
    }

    /// <summary>
    /// An Excel table definition.
    /// </summary>
    public sealed class TableDefinition
    {
        /// <summary>
        /// Gets the table name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the display name, or <see cref="Name"/> when not specified.
        /// </summary>
        public string DisplayName { get; }

        /// <summary>
        /// Gets the cell range the table covers.
        /// </summary>
        public string Ref { get; }

        /// <summary>
        /// Gets a value indicating whether the totals row is shown.
        /// </summary>
        public bool ShowTotalsRow { get; }

        /// <summary>
        /// Gets a value indicating whether the auto-filter buttons are shown.
        /// </summary>
        public bool ShowAutoFilter { get; }

        /// <summary>
        /// Gets the built-in table style. Overridden when a custom style name is set via <see cref="WithTableStyle(string)"/>.
        /// </summary>
        public TableStyle TableStyle { get; private set; }

        /// <summary>
        /// Gets the style name written to the OOXML. A custom string takes precedence; otherwise the <see cref="TableStyle"/> name is used.
        /// </summary>
        public string TableStyleName => _tableStyleRaw ?? TableStyle.ToString();

        private string? _tableStyleRaw;
        private List<TableColumnDefinition>? _columns;

        /// <summary>
        /// Gets the column definitions.
        /// </summary>
        public IReadOnlyList<TableColumnDefinition> Columns => _columns ??= new();

        /// <summary>
        /// Creates an Excel table definition.
        /// </summary>
        public TableDefinition(string name, string ref_, string? displayName = null,
            bool showTotalsRow = false, bool showAutoFilter = true,
            TableStyle tableStyle = TableStyle.TableStyleMedium2)
        {
            Name = name;
            Ref = ref_;
            DisplayName = displayName ?? name;
            ShowTotalsRow = showTotalsRow;
            ShowAutoFilter = showAutoFilter;
            TableStyle = tableStyle;
            _tableStyleRaw = null;
        }

        /// <summary>
        /// Adds a column. Returns this instance for chaining.
        /// </summary>
        public TableDefinition WithColumn(string name, string? totalsRowFunction = null)
        {
            (_columns ??= new()).Add(new TableColumnDefinition(name, totalsRowFunction));
            return this;
        }

        /// <summary>
        /// Sets a built-in table style. The value is compile-time validated, so prefer this overload when using a built-in style.
        /// </summary>
        public TableDefinition WithTableStyle(TableStyle tableStyle)
        {
            TableStyle = tableStyle;
            _tableStyleRaw = null;
            return this;
        }

        /// <summary>
        /// Sets a raw style name not included in <see cref="TableStyle"/>. The name is not validated, so a typo may cause Excel to ignore the style.
        /// </summary>
        public TableDefinition WithTableStyle(string tableStyle)
        {
            _tableStyleRaw = tableStyle;
            return this;
        }
    }
}
