using System;
using System.IO.Compression;

namespace Magicodes.IE.IO
{
    /// <summary>
    /// Controls compression and cell-reference behavior when writing an .xlsx workbook.
    /// </summary>
    public sealed class XlsxWriteOptions
    {
        /// <summary>
        /// The ZIP compression level. The default is <see cref="CompressionLevel.Fastest"/>.
        /// </summary>
        public CompressionLevel Compression { get; init; } = CompressionLevel.Fastest;

        /// <summary>
        /// Whether to write an A1 reference (for example, <c>A1</c>) on every cell.
        /// Disabling this yields a smaller file but reduces compatibility with some consumers.
        /// </summary>
        public bool StrictCellReferences { get; init; } = true;

        /// <summary>
        /// Creates an instance that sets only the compression level.
        /// </summary>
        public static XlsxWriteOptions WithCompression(CompressionLevel compression) => new() { Compression = compression };

        /// <summary>
        /// Creates an instance that omits cell references, for consumers that do not rely on them.
        /// </summary>
        public static XlsxWriteOptions WithoutCellReferences() => new() { StrictCellReferences = false };
    }
}
