using System;
using System.Buffers;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

namespace Magicodes.IE.IO
{
    /// <summary>
    /// Provides high-level, allocation-conscious APIs for reading and writing .xlsx workbooks without explicitly creating a writer or reader.
    /// </summary>
    /// <remarks>
    /// <para><b>Export:</b> use the <see cref="Write{T}(System.String, IEnumerable{T}, System.Action{ExportProfile{T}}, XlsxWriteOptions)"/> overloads to write rows to a file, stream, <see cref="IBufferWriter{T}"/>, or <see cref="byte"/> array.</para>
    /// <para><b>Import:</b> use the <see cref="Read{T}(Stream, XlsxReadOptions{T}?, Action{XlsxReadErrorInfo}?, bool)"/> overloads to stream deserialized rows from a workbook.</para>
    /// <para><b>Multi-sheet:</b> <see cref="WriteWorkbook(Stream, IReadOnlyList{SheetBase})"/>. <b>Template export:</b> <see cref="ExportByTemplateAsync{T}(System.String, System.String, T, CancellationToken)"/>.</para>
    /// <para>When no <see cref="ExportProfile{T}"/> is supplied, headers and formats are inferred from property names and from <c>[Display]</c>, <c>[Description]</c>, and <c>[DisplayFormat]</c> attributes. The underlying writer/reader is a dependency-free, streaming, low-allocation OOXML implementation and does not depend on EPPlus.</para>
    /// <para><b>Lifetime:</b> overloads that take a <c>path</c> own and dispose the underlying <see cref="FileStream"/>; overloads that take a <see cref="Stream"/> or <see cref="IBufferWriter{T}"/> do not, and the caller is responsible for disposal. Note: the Read/ReadAsync stream overloads DO own and dispose the supplied stream when enumeration completes.</para>
    /// </remarks>
    public static class Xlsx
    {
        /// <summary>
        /// Writes the specified rows to the .xlsx file at <paramref name="path"/>.
        /// </summary>
        /// <typeparam name="T">The row model type. May be a class, record, struct, or readonly struct.</typeparam>
        /// <param name="path">The output file path. Cannot be <see langword="null"/> or whitespace.</param>
        /// <param name="data">The rows to export. Any <see cref="IEnumerable{T}"/>.</param>
        /// <param name="configure">An optional callback that configures the header, columns, and styles through the fluent <see cref="ExportProfile{T}"/> API. When omitted, columns are inferred from the model.</param>
        /// <param name="options">Optional <see cref="XlsxWriteOptions"/>, such as the compression level and cell-reference strictness.</param>
        /// <exception cref="ArgumentNullException"><paramref name="path"/> or <paramref name="data"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="path"/> is empty or whitespace.</exception>
        public static void Write<T>(string path, IEnumerable<T> data, Action<ExportProfile<T>>? configure = null, XlsxWriteOptions? options = null)
        {
            ValidatePath(path);
            if (data is null) throw new ArgumentNullException(nameof(data));
            using var fs = File.Create(path);
            XlsxWritePipeline.Run(fs, data, configure, options);
        }

        /// <summary>
        /// Writes the specified rows to <paramref name="path"/> using a pre-built <see cref="ExportProfile{T}"/> instead of a fluent callback.
        /// </summary>
        /// <remarks>Use this overload when you reuse the same configuration across many exports, to avoid rebuilding the profile each time.</remarks>
        /// <exception cref="ArgumentNullException"><paramref name="path"/>, <paramref name="data"/>, or <paramref name="profile"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="path"/> is empty or whitespace.</exception>
        public static void Write<T>(string path, IEnumerable<T> data, ExportProfile<T> profile, XlsxWriteOptions? options = null)
        {
            ValidatePath(path);
            if (data is null) throw new ArgumentNullException(nameof(data));
            if (profile is null) throw new ArgumentNullException(nameof(profile));
            using var fs = File.Create(path);
            XlsxWritePipeline.Run(fs, data, profile, options);
        }

        /// <summary>
        /// Writes the specified rows to an existing stream. The caller owns and disposes <paramref name="output"/>.
        /// </summary>
        /// <remarks>Useful for writing directly to a response stream (for example, ASP.NET's <c>HttpResponse.Body</c>) or any externally managed stream.</remarks>
        /// <exception cref="ArgumentNullException"><paramref name="output"/> or <paramref name="data"/> is <see langword="null"/>.</exception>
        public static void Write<T>(Stream output, IEnumerable<T> data, Action<ExportProfile<T>>? configure = null, XlsxWriteOptions? options = null)
        {
            if (output is null) throw new ArgumentNullException(nameof(output));
            if (data is null) throw new ArgumentNullException(nameof(data));
            XlsxWritePipeline.Run(output, data, configure, options);
        }

        /// <summary>
        /// Writes the specified rows to an existing stream using a pre-built <see cref="ExportProfile{T}"/>. The caller owns and disposes <paramref name="output"/>.
        /// </summary>
        /// <exception cref="ArgumentNullException"><paramref name="output"/>, <paramref name="data"/>, or <paramref name="profile"/> is <see langword="null"/>.</exception>
        public static void Write<T>(Stream output, IEnumerable<T> data, ExportProfile<T> profile, XlsxWriteOptions? options = null)
        {
            if (output is null) throw new ArgumentNullException(nameof(output));
            if (data is null) throw new ArgumentNullException(nameof(data));
            if (profile is null) throw new ArgumentNullException(nameof(profile));
            XlsxWritePipeline.Run(output, data, profile, options);
        }

        /// <summary>
        /// Writes the specified rows to an <see cref="IBufferWriter{T}"/> (for example, a response buffer) for low-allocation output.
        /// </summary>
        /// <remarks>Avoids the extra <see cref="Stream"/> wrapper used by the stream overloads; bytes are written directly to the buffer.</remarks>
        /// <exception cref="ArgumentNullException"><paramref name="output"/> or <paramref name="data"/> is <see langword="null"/>.</exception>
        public static void Write<T>(IBufferWriter<byte> output, IEnumerable<T> data, Action<ExportProfile<T>>? configure = null, XlsxWriteOptions? options = null)
        {
            if (output is null) throw new ArgumentNullException(nameof(output));
            if (data is null) throw new ArgumentNullException(nameof(data));
            XlsxWritePipeline.Run(new BufferWriterStream(output), data, configure, options);
        }

        /// <summary>
        /// Writes the specified rows to an <see cref="IBufferWriter{T}"/> using a pre-built <see cref="ExportProfile{T}"/>.
        /// </summary>
        /// <exception cref="ArgumentNullException"><paramref name="output"/>, <paramref name="data"/>, or <paramref name="profile"/> is <see langword="null"/>.</exception>
        public static void Write<T>(IBufferWriter<byte> output, IEnumerable<T> data, ExportProfile<T> profile, XlsxWriteOptions? options = null)
        {
            if (output is null) throw new ArgumentNullException(nameof(output));
            if (data is null) throw new ArgumentNullException(nameof(data));
            if (profile is null) throw new ArgumentNullException(nameof(profile));
            XlsxWritePipeline.Run(new BufferWriterStream(output), data, profile, options);
        }

        /// <summary>
        /// Streams rows from an <see cref="IAsyncEnumerable{T}"/> to a stream, enumerating and writing concurrently so the data is never fully materialized.
        /// </summary>
        /// <remarks>The caller owns and disposes <paramref name="output"/>. Use this overload for large or paged data sources (for example, a database query). Unlike the <see cref="IEnumerable{T}"/> async overloads, the data source itself is asynchronous.</remarks>
        /// <exception cref="ArgumentNullException"><paramref name="output"/> or <paramref name="data"/> is <see langword="null"/>.</exception>
        public static async Task WriteAsync<T>(Stream output, IAsyncEnumerable<T> data, Action<ExportProfile<T>>? configure = null, XlsxWriteOptions? options = null, CancellationToken cancellationToken = default)
        {
            if (output is null) throw new ArgumentNullException(nameof(output));
            if (data is null) throw new ArgumentNullException(nameof(data));
            var compression = options?.Compression ?? System.IO.Compression.CompressionLevel.Fastest;
            var strictCellReferences = options?.StrictCellReferences ?? true;
            await using var writer = new XlsxWriter(output, sheetName: null, compression, defaultRowHeight: 0, strictCellReferences);
            await XlsxWritePipeline.RunAsync(writer, data, configure, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Streams rows from an <see cref="IAsyncEnumerable{T}"/> to a stream using a pre-built <see cref="ExportProfile{T}"/>.
        /// </summary>
        /// <exception cref="ArgumentNullException"><paramref name="output"/>, <paramref name="data"/>, or <paramref name="profile"/> is <see langword="null"/>.</exception>
        public static async Task WriteAsync<T>(Stream output, IAsyncEnumerable<T> data, ExportProfile<T> profile, XlsxWriteOptions? options = null, CancellationToken cancellationToken = default)
        {
            if (output is null) throw new ArgumentNullException(nameof(output));
            if (data is null) throw new ArgumentNullException(nameof(data));
            if (profile is null) throw new ArgumentNullException(nameof(profile));
            var compression = options?.Compression ?? System.IO.Compression.CompressionLevel.Fastest;
            var strictCellReferences = options?.StrictCellReferences ?? true;
            await using var writer = new XlsxWriter(output, sheetName: null, compression, defaultRowHeight: 0, strictCellReferences);
            await XlsxWritePipeline.RunAsync(writer, data, profile, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Writes an in-memory <see cref="IEnumerable{T}"/> to a stream asynchronously.
        /// </summary>
        /// <remarks>The data source is not asynchronous here; only the writing is. For an asynchronous data source, use the <see cref="IAsyncEnumerable{T}"/> overload.</remarks>
        /// <exception cref="ArgumentNullException"><paramref name="output"/> or <paramref name="data"/> is <see langword="null"/>.</exception>
        public static async Task WriteAsync<T>(Stream output, IEnumerable<T> data, Action<ExportProfile<T>>? configure = null, XlsxWriteOptions? options = null, CancellationToken cancellationToken = default)
        {
            if (output is null) throw new ArgumentNullException(nameof(output));
            if (data is null) throw new ArgumentNullException(nameof(data));
            var compression = options?.Compression ?? System.IO.Compression.CompressionLevel.Fastest;
            var strictCellReferences = options?.StrictCellReferences ?? true;
            await using var writer = new XlsxWriter(output, sheetName: null, compression, defaultRowHeight: 0, strictCellReferences);
            await XlsxWritePipeline.RunAsync(writer, data, configure, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Writes an in-memory <see cref="IEnumerable{T}"/> to a stream asynchronously using a pre-built <see cref="ExportProfile{T}"/>.
        /// </summary>
        /// <exception cref="ArgumentNullException"><paramref name="output"/>, <paramref name="data"/>, or <paramref name="profile"/> is <see langword="null"/>.</exception>
        public static async Task WriteAsync<T>(Stream output, IEnumerable<T> data, ExportProfile<T> profile, XlsxWriteOptions? options = null, CancellationToken cancellationToken = default)
        {
            if (output is null) throw new ArgumentNullException(nameof(output));
            if (data is null) throw new ArgumentNullException(nameof(data));
            if (profile is null) throw new ArgumentNullException(nameof(profile));
            var compression = options?.Compression ?? System.IO.Compression.CompressionLevel.Fastest;
            var strictCellReferences = options?.StrictCellReferences ?? true;
            await using var writer = new XlsxWriter(output, sheetName: null, compression, defaultRowHeight: 0, strictCellReferences);
            await XlsxWritePipeline.RunAsync(writer, data, profile, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Asynchronously writes the specified rows to the .xlsx file at <paramref name="path"/>.
        /// </summary>
        /// <remarks>Equivalent to opening a <see cref="FileStream"/> and calling the stream overload; the file stream is owned and disposed by this method.</remarks>
        public static async Task WriteAsync<T>(string path, IEnumerable<T> data, Action<ExportProfile<T>>? configure = null, XlsxWriteOptions? options = null, CancellationToken cancellationToken = default)
        {
            ValidatePath(path);
            if (data is null) throw new ArgumentNullException(nameof(data));
            var compression = options?.Compression ?? System.IO.Compression.CompressionLevel.Fastest;
            var strictCellReferences = options?.StrictCellReferences ?? true;
            using var fs = File.Create(path);
            await using var writer = new XlsxWriter(fs, sheetName: null, compression, defaultRowHeight: 0, strictCellReferences);
            await XlsxWritePipeline.RunAsync(writer, data, configure, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Asynchronously writes the specified rows to the .xlsx file at <paramref name="path"/> using a pre-built profile.
        /// </summary>
        public static async Task WriteAsync<T>(string path, IEnumerable<T> data, ExportProfile<T> profile, XlsxWriteOptions? options = null, CancellationToken cancellationToken = default)
        {
            ValidatePath(path);
            if (data is null) throw new ArgumentNullException(nameof(data));
            if (profile is null) throw new ArgumentNullException(nameof(profile));
            var compression = options?.Compression ?? System.IO.Compression.CompressionLevel.Fastest;
            var strictCellReferences = options?.StrictCellReferences ?? true;
            using var fs = File.Create(path);
            await using var writer = new XlsxWriter(fs, sheetName: null, compression, defaultRowHeight: 0, strictCellReferences);
            await XlsxWritePipeline.RunAsync(writer, data, profile, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Asynchronously streams rows from an <see cref="IAsyncEnumerable{T}"/> to the .xlsx file at <paramref name="path"/>.
        /// </summary>
        public static async Task WriteAsync<T>(string path, IAsyncEnumerable<T> data, Action<ExportProfile<T>>? configure = null, XlsxWriteOptions? options = null, CancellationToken cancellationToken = default)
        {
            ValidatePath(path);
            if (data is null) throw new ArgumentNullException(nameof(data));
            var compression = options?.Compression ?? System.IO.Compression.CompressionLevel.Fastest;
            var strictCellReferences = options?.StrictCellReferences ?? true;
            using var fs = File.Create(path);
            await using var writer = new XlsxWriter(fs, sheetName: null, compression, defaultRowHeight: 0, strictCellReferences);
            await XlsxWritePipeline.RunAsync(writer, data, configure, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Asynchronously streams rows from an <see cref="IAsyncEnumerable{T}"/> to the .xlsx file at <paramref name="path"/> using a pre-built profile.
        /// </summary>
        public static async Task WriteAsync<T>(string path, IAsyncEnumerable<T> data, ExportProfile<T> profile, XlsxWriteOptions? options = null, CancellationToken cancellationToken = default)
        {
            ValidatePath(path);
            if (data is null) throw new ArgumentNullException(nameof(data));
            if (profile is null) throw new ArgumentNullException(nameof(profile));
            var compression = options?.Compression ?? System.IO.Compression.CompressionLevel.Fastest;
            var strictCellReferences = options?.StrictCellReferences ?? true;
            using var fs = File.Create(path);
            await using var writer = new XlsxWriter(fs, sheetName: null, compression, defaultRowHeight: 0, strictCellReferences);
            await XlsxWritePipeline.RunAsync(writer, data, profile, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Exports the specified rows to a <see cref="byte"/> array.
        /// </summary>
        /// <remarks>For very large datasets, prefer the stream overloads to avoid holding the entire workbook in memory. The initial <see cref="MemoryStream"/> capacity is estimated from the collection size to reduce reallocations.</remarks>
        /// <exception cref="ArgumentNullException"><paramref name="data"/> is <see langword="null"/>.</exception>
        public static byte[] ToBytes<T>(IEnumerable<T> data, Action<ExportProfile<T>>? configure = null, XlsxWriteOptions? options = null)
        {
            if (data is null) throw new ArgumentNullException(nameof(data));
            using var ms = new MemoryStream(EstimateToBytesCapacity(data));
            XlsxWritePipeline.Run(ms, data, configure, options);
            return ms.ToArray();
        }

        /// <summary>
        /// Exports the specified rows to a <see cref="byte"/> array using a pre-built <see cref="ExportProfile{T}"/>.
        /// </summary>
        /// <exception cref="ArgumentNullException"><paramref name="data"/> or <paramref name="profile"/> is <see langword="null"/>.</exception>
        public static byte[] ToBytes<T>(IEnumerable<T> data, ExportProfile<T> profile, XlsxWriteOptions? options = null)
        {
            if (data is null) throw new ArgumentNullException(nameof(data));
            if (profile is null) throw new ArgumentNullException(nameof(profile));
            using var ms = new MemoryStream(EstimateToBytesCapacity(data));
            XlsxWritePipeline.Run(ms, data, profile, options);
            return ms.ToArray();
        }

        // Only pre-size ToBytes when the source is a large, countable collection, to avoid
        // over-allocating for small or deferred sequences.
        private static int EstimateToBytesCapacity<T>(IEnumerable<T> data)
        {
            if (data is ICollection<T> coll && coll.Count > 1024)
                return Math.Min(coll.Count * 64, 64 * 1024 * 1024);
            return 0;
        }

        /// <summary>
        /// Reads the first worksheet of a .xlsx workbook and lazily returns the deserialized rows as <typeparamref name="T"/>.
        /// </summary>
        /// <typeparam name="T">The row model type. Must have a parameterless constructor.</typeparam>
        /// <param name="stream">The stream containing the .xlsx data. It is disposed when enumeration completes.</param>
        /// <param name="profile">Optional <see cref="XlsxReadOptions{T}"/> that configures header mapping and date handling.</param>
        /// <param name="onParseError">Optional callback invoked when a cell cannot be parsed. When omitted, a <see cref="XlsxException"/> is thrown on the first parse error.</param>
        /// <returns>A lazy <see cref="IEnumerable{T}"/>. Each row is parsed on demand as you iterate; do not enumerate the result more than once.</returns>
        /// <remarks>When a source-generated reader is available for <typeparamref name="T"/>, it is used for reflection-free reading; otherwise reflection is used.</remarks>
        public static IEnumerable<T> Read<T>(Stream stream, XlsxReadOptions<T>? profile = null, Action<XlsxReadErrorInfo>? onParseError = null, bool leaveOpen = false) where T : new()
        {
            using var reader = new XlsxReader(stream, leaveOpen);
            var headers = reader.ReadHeader();
            var converters = profile?.GetConverters();
            int rowIndex = 0;

            var generated = XlsxGeneratedTypeMetadataRegistry.TryGet<T>();
            if (generated is not null)
            {
                var generatedResolver = XlsxReadPipeline.BuildGeneratedResolver(headers, profile, generated);
                while (true)
                {
                    var row = reader.ReadNextRowView();
                    if (row is null) yield break;
                    var item = new T();
                    for (int i = 0; i < row.Count; i++)
                    {
                        var cell = row[i];
                        if (cell is null) continue;
                        var property = generatedResolver(i);
                        if (property is null) continue;
                        XlsxReadPipeline.SetGeneratedProperty(item, property, cell, onParseError, rowIndex, i, headers.Length > i ? headers[i] : null, converters, reader.Date1904);
                    }
                    rowIndex++;
                    yield return item;
                }
            }

            var resolver = XlsxReadPipeline.BuildResolver(headers, profile);
            while (true)
            {
                var row = reader.ReadNextRowView();
                if (row is null) yield break;
                object item = new T();
                for (int i = 0; i < row.Count; i++)
                {
                    var cell = row[i];
                    if (cell is null) continue;
                    var prop = resolver(i);
                    if (prop is null) continue;
                    XlsxReadPipeline.SetCellProperty(item, prop, cell, onParseError, rowIndex, i, headers.Length > i ? headers[i] : null, converters, reader.Date1904);
                }
                rowIndex++;
                yield return (T)item;
            }
        }

        /// <summary>
        /// Reads the first worksheet of the .xlsx file at <paramref name="path"/> and lazily returns the deserialized rows as <typeparamref name="T"/>.
        /// </summary>
        /// <remarks>This overload owns and disposes the underlying <see cref="FileStream"/>; the file is read only while the returned sequence is enumerated.</remarks>
        public static IEnumerable<T> Read<T>(string path, XlsxReadOptions<T>? profile = null, Action<XlsxReadErrorInfo>? onParseError = null) where T : new()
        {
            ValidatePath(path);
            using var fs = File.OpenRead(path);
            foreach (var item in Read<T>(fs, profile, onParseError))
                yield return item;
        }

        /// <summary>
        /// Reads the first worksheet of a .xlsx workbook asynchronously and returns the deserialized rows as <typeparamref name="T"/>.
        /// </summary>
        /// <remarks>Supports cancellation via <paramref name="cancellationToken"/>. Otherwise the behavior matches <see cref="Read{T}(Stream, XlsxReadOptions{T}?, Action{XlsxReadErrorInfo}?, bool)"/>.</remarks>
        /// <returns>A lazy <see cref="IAsyncEnumerable{T}"/>. Do not enumerate the result more than once.</returns>
        public static async IAsyncEnumerable<T> ReadAsync<T>(Stream stream, XlsxReadOptions<T>? profile = null, Action<XlsxReadErrorInfo>? onParseError = null, [EnumeratorCancellation] CancellationToken cancellationToken = default, bool leaveOpen = false) where T : new()
        {
            cancellationToken.ThrowIfCancellationRequested();
            using var reader = new XlsxReader(stream, leaveOpen);
            var headers = await reader.ReadHeaderAsync(cancellationToken).ConfigureAwait(false);
            var converters = profile?.GetConverters();
            int rowIndex = 0;

            var generated = XlsxGeneratedTypeMetadataRegistry.TryGet<T>();
            if (generated is not null)
            {
                var generatedResolver = XlsxReadPipeline.BuildGeneratedResolver(headers, profile, generated);
                while (!cancellationToken.IsCancellationRequested)
                {
                    var row = await reader.ReadNextRowViewAsync(cancellationToken).ConfigureAwait(false);
                    if (row is null) yield break;
                    var item = new T();
                    for (int i = 0; i < row.Count; i++)
                    {
                        var cell = row[i];
                        if (cell is null) continue;
                        var property = generatedResolver(i);
                        if (property is null) continue;
                        XlsxReadPipeline.SetGeneratedProperty(item, property, cell, onParseError, rowIndex, i, headers.Length > i ? headers[i] : null, converters, reader.Date1904);
                    }
                    rowIndex++;
                    yield return item;
                }
                cancellationToken.ThrowIfCancellationRequested();
                yield break;
            }

            var resolver = XlsxReadPipeline.BuildResolver(headers, profile);
            while (!cancellationToken.IsCancellationRequested)
            {
                var row = await reader.ReadNextRowViewAsync(cancellationToken).ConfigureAwait(false);
                if (row is null) yield break;
                object item = new T();
                for (int i = 0; i < row.Count; i++)
                {
                    var cell = row[i];
                    if (cell is null) continue;
                    var prop = resolver(i);
                    if (prop is null) continue;
                    XlsxReadPipeline.SetCellProperty(item, prop, cell, onParseError, rowIndex, i, headers.Length > i ? headers[i] : null, converters, reader.Date1904);
                }
                rowIndex++;
                yield return (T)item;
            }
            cancellationToken.ThrowIfCancellationRequested();
        }

        /// <summary>
        /// Asynchronously reads the first worksheet of the .xlsx file at <paramref name="path"/> and lazily returns the deserialized rows as <typeparamref name="T"/>.
        /// </summary>
        /// <remarks>This overload owns and disposes the underlying <see cref="FileStream"/>; the file is read only while the returned sequence is enumerated.</remarks>
        public static async IAsyncEnumerable<T> ReadAsync<T>(string path, XlsxReadOptions<T>? profile = null, Action<XlsxReadErrorInfo>? onParseError = null, [EnumeratorCancellation] CancellationToken cancellationToken = default) where T : new()
        {
            ValidatePath(path);
            using var fs = File.OpenRead(path);
            await foreach (var item in ReadAsync<T>(fs, profile, onParseError, cancellationToken: cancellationToken).ConfigureAwait(false))
                yield return item;
        }

        /// <summary>
        /// Writes multiple sheets to a stream. At least one non-<see langword="null"/> sheet must be supplied.
        /// </summary>
        /// <remarks>Each sheet's type and configuration are determined by its <see cref="SheetBase"/> subclass. The caller owns and disposes <paramref name="output"/>.</remarks>
        /// <exception cref="ArgumentNullException"><paramref name="output"/> or <paramref name="sheets"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="sheets"/> is empty or contains a <see langword="null"/> element.</exception>
        public static void WriteWorkbook(Stream output, IReadOnlyList<SheetBase> sheets)
        {
            if (output is null) throw new ArgumentNullException(nameof(output));
            if (sheets is null) throw new ArgumentNullException(nameof(sheets));
            if (sheets.Count == 0) throw new ArgumentException("At least one sheet is required.", nameof(sheets));
            using var writer = new XlsxWriter(output);
            for (int i = 0; i < sheets.Count; i++)
            {
                if (sheets[i] is null) throw new ArgumentException("Sheet cannot be null.", nameof(sheets));
                XlsxWritePipeline.BuildAndWrite(writer, sheets[i]);
            }
        }

        /// <summary>
        /// Writes multiple sheets to a stream, supplied as a parameter list.
        /// </summary>
        /// <remarks>Equivalent to wrapping the sheets in an <see cref="IReadOnlyList{T}"/> and calling <see cref="WriteWorkbook(Stream, IReadOnlyList{SheetBase})"/>.</remarks>
        public static void WriteWorkbook(Stream output, params SheetBase[] sheets)
        {
            if (sheets is null) throw new ArgumentNullException(nameof(sheets));
            WriteWorkbook(output, (IReadOnlyList<SheetBase>)sheets);
        }

        /// <summary>
        /// Writes multiple sheets to an <see cref="IBufferWriter{T}"/>. At least one non-<see langword="null"/> sheet must be supplied.
        /// </summary>
        /// <exception cref="ArgumentNullException"><paramref name="output"/> or <paramref name="sheets"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="sheets"/> is empty or contains a <see langword="null"/> element.</exception>
        public static void WriteWorkbook(IBufferWriter<byte> output, IReadOnlyList<SheetBase> sheets)
        {
            if (output is null) throw new ArgumentNullException(nameof(output));
            if (sheets is null) throw new ArgumentNullException(nameof(sheets));
            if (sheets.Count == 0) throw new ArgumentException("At least one sheet is required.", nameof(sheets));
            using var writer = new XlsxWriter(new BufferWriterStream(output));
            for (int i = 0; i < sheets.Count; i++)
            {
                if (sheets[i] is null) throw new ArgumentException("Sheet cannot be null.", nameof(sheets));
                XlsxWritePipeline.BuildAndWrite(writer, sheets[i]);
            }
        }

        /// <summary>
        /// Writes multiple sheets to an <see cref="IBufferWriter{T}"/>, supplied as a parameter list.
        /// </summary>
        public static void WriteWorkbook(IBufferWriter<byte> output, params SheetBase[] sheets)
        {
            if (sheets is null) throw new ArgumentNullException(nameof(sheets));
            WriteWorkbook(output, (IReadOnlyList<SheetBase>)sheets);
        }

        /// <summary>
        /// Writes multiple sheets and returns them as a <see cref="byte"/> array.
        /// </summary>
        /// <exception cref="ArgumentNullException"><paramref name="sheets"/> is <see langword="null"/>.</exception>
        /// <exception cref="ArgumentException"><paramref name="sheets"/> is empty or contains a <see langword="null"/> element.</exception>
        public static byte[] WriteWorkbookToBytes(IReadOnlyList<SheetBase> sheets)
        {
            if (sheets is null) throw new ArgumentNullException(nameof(sheets));
            if (sheets.Count == 0) throw new ArgumentException("At least one sheet is required.", nameof(sheets));
            using var ms = new MemoryStream();
            using (var writer = new XlsxWriter(ms))
            {
                for (int i = 0; i < sheets.Count; i++)
                {
                    if (sheets[i] is null) throw new ArgumentException("Sheet cannot be null.", nameof(sheets));
                    XlsxWritePipeline.BuildAndWrite(writer, sheets[i]);
                }
            }
            return ms.ToArray();
        }

        /// <summary>
        /// Writes multiple sheets and returns them as a <see cref="byte"/> array, supplied as a parameter list.
        /// </summary>
        public static byte[] WriteWorkbookToBytes(params SheetBase[] sheets)
        {
            if (sheets is null) throw new ArgumentNullException(nameof(sheets));
            return WriteWorkbookToBytes((IReadOnlyList<SheetBase>)sheets);
        }

        /// <summary>
        /// Replaces <c>{{PropertyName}}</c> placeholders in an .xlsx template with values from <paramref name="data"/> and writes the result to <paramref name="outputPath"/>.
        /// </summary>
        /// <remarks>The template is opened and closed by this method, and the output file is created and closed as well.</remarks>
        public static Task ExportByTemplateAsync<T>(string templatePath, string outputPath, T data, CancellationToken cancellationToken = default)
            => XlsxTemplateExporter.ExportAsync(templatePath, outputPath, data, cancellationToken);

        /// <summary>
        /// Exports a model into a template stream and writes the result to an output stream. Both streams are owned and disposed by the caller.
        /// </summary>
        /// <remarks>Placeholders take the form <c>{{PropertyName}}</c>; the lifetimes of both streams are the caller's responsibility.</remarks>
        public static Task ExportByTemplateAsync<T>(Stream templateStream, Stream outputStream, T data, CancellationToken cancellationToken = default)
            => XlsxTemplateExporter.ExportAsync(templateStream, outputStream, data, cancellationToken);

        /// <summary>
        /// Writes the given untyped data to <paramref name="sheetName"/> using an already-open <see cref="XlsxWriter"/>.
        /// </summary>
        internal static void WriteSheetUntyped(XlsxWriter writer, string sheetName, System.Collections.IEnumerable data)
            => XlsxWritePipeline.WriteSheetUntyped(writer, sheetName, data);

        /// <summary>
        /// Writes the given typed data to <paramref name="sheetName"/> using an already-open <see cref="XlsxWriter"/>.
        /// </summary>
        internal static void WriteSheet<T>(XlsxWriter writer, string sheetName, IEnumerable<T> data, ExportProfile<T>? profile)
            => XlsxWritePipeline.WriteSheet(writer, sheetName, data, profile);

        // Throws if the output path is null or whitespace.
        private static void ValidatePath(string path)
        {
            if (path is null) throw new ArgumentNullException(nameof(path));
            if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Path cannot be empty.", nameof(path));
        }
    }
}
