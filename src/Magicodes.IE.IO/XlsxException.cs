using System;
using System.ComponentModel;

namespace Magicodes.IE.IO
{
    /// <summary>
    /// The exception that is thrown when an error occurs while reading or writing an .xlsx workbook, optionally carrying the cell location of the failure.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class XlsxException : Exception
    {
        /// <summary>
        /// Gets the zero-based row index where the error occurred, or <see langword="null"/> if not applicable.
        /// </summary>
        public int? RowIndex { get; }

        /// <summary>
        /// Gets the zero-based column index where the error occurred, or <see langword="null"/> if not applicable.
        /// </summary>
        public int? ColIndex { get; }

        /// <summary>
        /// Gets the A1-style cell reference (for example, <c>A2</c>), or <see langword="null"/> if not available.
        /// </summary>
        public string? CellRef { get; }

        /// <summary>
        /// Gets the header text of the failing column, or <see langword="null"/> if not available.
        /// </summary>
        public string? Header { get; }

        /// <summary>
        /// Gets the target property name, or <see langword="null"/> if not available.
        /// </summary>
        public string? PropertyName { get; }

        /// <summary>
        /// Gets the raw cell text that could not be parsed, or <see langword="null"/> if not available.
        /// </summary>
        public string? RawCellValue { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="XlsxException"/> class with a specified error message.
        /// </summary>
        public XlsxException(string message) : base(message) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="XlsxException"/> class with a specified error message and inner exception.
        /// </summary>
        public XlsxException(string message, Exception innerException) : base(message, innerException) { }

        /// <summary>
        /// Initializes a new instance of the <see cref="XlsxException"/> class with a message and the cell location of the failure.
        /// </summary>
        public XlsxException(string message, int? rowIndex = null, int? colIndex = null,
            string? cellRef = null, string? header = null, string? propertyName = null, string? rawCellValue = null)
            : base(message)
        {
            RowIndex = rowIndex;
            ColIndex = colIndex;
            CellRef = cellRef;
            Header = header;
            PropertyName = propertyName;
            RawCellValue = rawCellValue;
        }

        /// <summary>
        /// Gets the error message, appended with any available cell location details.
        /// </summary>
        public override string Message
        {
            get
            {
                var baseMsg = base.Message;
                var loc = new System.Text.StringBuilder(128);
                if (RowIndex.HasValue) loc.Append("Row=").Append(RowIndex.Value).Append(' ');
                if (ColIndex.HasValue) loc.Append("Col=").Append(ColIndex.Value).Append(' ');
                if (CellRef is not null) loc.Append("CellRef=").Append(CellRef).Append(' ');
                if (Header is not null) loc.Append("Header=").Append(Header).Append(' ');
                if (PropertyName is not null) loc.Append("Property=").Append(PropertyName).Append(' ');
                if (RawCellValue is not null) loc.Append("RawValue=").Append(RawCellValue).Append(' ');
                return loc.Length > 0
                    ? $"{baseMsg} (at {loc.ToString().TrimEnd()})"
                    : baseMsg;
            }
        }
    }
}
