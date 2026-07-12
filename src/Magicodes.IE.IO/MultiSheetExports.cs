using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;

namespace Magicodes.IE.IO
{
    /// <summary>
    /// Base class for the sheets that can be written by <c>Xlsx.WriteWorkbook(Stream, IReadOnlyList&lt;SheetBase&gt;)</c>.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public abstract class SheetBase
    {
        /// <summary>
        /// Gets the name of the sheet.
        /// </summary>
        public string SheetName { get; }

        protected SheetBase(string sheetName)
        {
            SheetName = sheetName ?? throw new ArgumentNullException(nameof(sheetName));
        }

        internal abstract void WriteTo(XlsxWriter writer);
    }

    /// <summary>
    /// A worksheet whose rows are supplied as a non-generic <see cref="IEnumerable"/> and exported with default settings.
    /// </summary>
    public sealed class Sheet : SheetBase
    {
        /// <summary>
        /// Gets the row data for the sheet.
        /// </summary>
        public IEnumerable Data { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Sheet"/> class with the default export configuration.
        /// </summary>
        public Sheet(string sheetName, IEnumerable data)
            : base(sheetName)
        {
            Data = data ?? throw new ArgumentNullException(nameof(data));
        }

        internal override void WriteTo(XlsxWriter writer)
        {
            if (writer is null) throw new ArgumentNullException(nameof(writer));
            Xlsx.WriteSheetUntyped(writer, SheetName, Data);
        }
    }

    /// <summary>
    /// A strongly typed worksheet whose rows are exported using an optional <see cref="ExportProfile{T}"/>.
    /// </summary>
    /// <typeparam name="T">The row model type.</typeparam>
    public sealed class Sheet<T> : SheetBase
    {
        private readonly ExportProfile<T>? _profile;

        /// <summary>
        /// Gets the row data for the sheet.
        /// </summary>
        public IEnumerable<T> Data { get; }

        /// <summary>
        /// Initializes a new instance of the <see cref="Sheet{T}"/> class with the default export configuration.
        /// </summary>
        public Sheet(string sheetName, IEnumerable<T> data)
            : base(sheetName)
        {
            Data = data ?? throw new ArgumentNullException(nameof(data));
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Sheet{T}"/> class with a specific export configuration.
        /// </summary>
        public Sheet(string sheetName, IEnumerable<T> data, ExportProfile<T> profile)
            : this(sheetName, data)
        {
            _profile = profile ?? throw new ArgumentNullException(nameof(profile));
        }

        internal override void WriteTo(XlsxWriter writer)
        {
            if (writer is null) throw new ArgumentNullException(nameof(writer));
            Xlsx.WriteSheet(writer, SheetName, Data, _profile);
        }
    }
}
