using System;
using System.ComponentModel;

namespace Magicodes.IE.IO
{
    /// <summary>
    /// An intermediate cell value produced by the writer and converters.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public readonly struct CellValue
    {
        /// <summary>
        /// Gets the kind of value this cell holds.
        /// </summary>
        public readonly CellType Type;
        /// <summary>
        /// Gets the string value. Valid when <see cref="Type"/> is <see cref="CellType.String"/>.
        /// </summary>
        public readonly string? StringValue;
        /// <summary>
        /// Gets the numeric value. Valid when <see cref="Type"/> is <see cref="CellType.Number"/>.
        /// </summary>
        public readonly double NumberValue;
        /// <summary>
        /// Gets the Boolean value. Valid when <see cref="Type"/> is <see cref="CellType.Boolean"/>.
        /// </summary>
        public readonly bool BoolValue;

        private CellValue(CellType type, string? s, double n, bool b)
        {
            Type = type; StringValue = s; NumberValue = n; BoolValue = b;
        }

        /// <summary>
        /// Represents an empty cell.
        /// </summary>
        public static CellValue Null => new(CellType.Null, null, 0, false);
        /// <summary>
        /// Creates a string cell, or <see cref="Null"/> when <paramref name="s"/> is <see langword="null"/>.
        /// </summary>
        public static CellValue FromString(string? s) => s is null ? Null : new(CellType.String, s, 0, false);
        /// <summary>
        /// Creates a numeric cell.
        /// </summary>
        public static CellValue FromNumber(double d) => new(CellType.Number, null, d, false);
        /// <summary>
        /// Creates a numeric cell from a 64-bit integer.
        /// </summary>
        public static CellValue FromInteger(long l) => new(CellType.Number, null, l, false);
        /// <summary>
        /// Creates a date cell using the Excel OLE Automation date serial number.
        /// </summary>
        public static CellValue FromDateTime(DateTime dt) => new(CellType.Number, null, dt.ToOADate(), false);
        /// <summary>
        /// Creates a Boolean cell.
        /// </summary>
        public static CellValue FromBool(bool b) => new(CellType.Boolean, null, 0, b);
    }

    /// <summary>
    /// The kind of value stored in a <see cref="CellValue"/>.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public enum CellType : byte
    {
        /// <summary>
        /// An empty cell.
        /// </summary>
        Null = 0,
        /// <summary>
        /// A string value.
        /// </summary>
        String = 1,
        /// <summary>
        /// A numeric value.
        /// </summary>
        Number = 2,
        /// <summary>
        /// A Boolean value.
        /// </summary>
        Boolean = 3,
        /// <summary>
        /// A formula.
        /// </summary>
        Formula = 4,
    }
}
