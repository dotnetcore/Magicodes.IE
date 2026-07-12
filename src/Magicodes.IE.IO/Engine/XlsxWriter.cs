
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.ComponentModel;

namespace Magicodes.IE.IO
{

    /// <summary>
    /// Advanced streaming writer for xlsx workbooks.
    /// For ordinary exports, use <see cref="Xlsx"/>.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed partial class XlsxWriter : IDisposable, IAsyncDisposable
    {
        private readonly ForwardOnlyZipWriter _zip;
        private readonly ByteBufferWriter _sink;
        private Stream? _sheetStream;
        private int _rowIndex;
        private readonly byte[] _rowNumberBuf = new byte[12];
        private int _rowNumberLen;
        private ReadOnlySpan<byte> RowNumberRef => _rowNumberBuf.AsSpan(0, _rowNumberLen);
        private int _currentColIndex;
        private bool _disposed;
        private IReadOnlyList<string> _numFmts = Array.Empty<string>();
        private ColumnMeta[]? _currentColumns;
        private readonly List<string> _sheetNames = new();
        private int _currentSheetIndex = -1;
        private CompressionLevel _compression = CompressionLevel.Fastest;
        private double? _defaultRowHeight;
        private bool _strictCellReferences = true;
        private bool _sheetDataStarted;
        private byte[][]? _colLetterCache;
        private readonly List<SheetState> _sheets = new();
        private readonly List<(string Path, byte[] Bytes)> _imageBytes = new();
        private int _imageCounter = 0;
        private double? _nextRowHeight;
        private bool _useSharedStrings;
        private readonly Dictionary<string, int> _sharedStringLookup = new();
        private readonly List<string> _sharedStrings = new();
        private int _sstRefCount;
        private readonly List<NamedRange> _namedRanges = new();

        /// <summary>
        /// Initializes a new writer. The caller owns and disposes the output stream. When <paramref name="sheetName"/> is supplied, the first sheet is created immediately.
        /// </summary>
        public XlsxWriter(Stream output, string? sheetName = null, CompressionLevel compression = CompressionLevel.Fastest, double defaultRowHeight = 0, bool strictCellReferences = true)
        {
            if (output is null) throw new ArgumentNullException(nameof(output));

            _zip = new ForwardOnlyZipWriter(output, leaveOpen: true);
            _sink = new ByteBufferWriter(initialSize: SheetStreamBufferBytes);
            _compression = compression;
            _strictCellReferences = strictCellReferences;
            if (defaultRowHeight > 0) _defaultRowHeight = defaultRowHeight;

            try
            {
                WriteRels();

                if (sheetName is not null)
                    AddSheet(sheetName);
            }
            catch
            {
                _zip.Dispose();
                _sink.Dispose();
                throw;
            }
        }

        /// <summary>
        /// Gets the one-based row index of the current sheet.
        /// </summary>
        public int CurrentRow => _rowIndex;

        /// <summary>
        /// Gets the number of sheets created so far.
        /// </summary>
        public int SheetCount => _sheetNames.Count;

        /// <summary>
        /// Gets the names of the sheets created so far.
        /// </summary>
        public IReadOnlyList<string> SheetNames => _sheetNames;

        internal bool SupportsAutoSharedStrings => _compression != CompressionLevel.NoCompression;

        /// <summary>
        /// Registers the numeric formats used in the workbook.
        /// </summary>
        public void SetNumFmts(IReadOnlyList<string> numFmts)
        {
            EnsureNotDisposed();
            var existing = _numFmts.ToList();
            foreach (var fmt in numFmts ?? Array.Empty<string>())
            {
                if (!string.IsNullOrEmpty(fmt) && !existing.Contains(fmt))
                    existing.Add(fmt);
            }
            _numFmts = existing;
        }

        /// <summary>
        /// Sets the default row height used by subsequent sheets.
        /// </summary>
        public void SetDefaultRowHeight(double height)
        {
            EnsureNotDisposed();
            if (height <= 0 || double.IsNaN(height)) throw new ArgumentOutOfRangeException(nameof(height), "Row height must be greater than 0.");
            _defaultRowHeight = height;
        }

        /// <summary>
        /// Sets the row height for the next row only; it reverts to the default after that row is written.
        /// </summary>
        public void SetNextRowHeight(double height)
        {
            EnsureNotDisposed();
            if (height <= 0 || double.IsNaN(height)) throw new ArgumentOutOfRangeException(nameof(height), "Row height must be greater than 0.");
            _nextRowHeight = height;
        }

        /// <summary>
        /// Enables the shared string table. Call this before writing the header or any data.
        /// </summary>
        public void EnableSharedStrings()
        {
            EnsureNotDisposed();
            _useSharedStrings = true;
        }

        /// <summary>
        /// Builds the style pool for the current sheet from the supplied column metadata.
        /// </summary>
        public void ResolveColumnStyles(ColumnMeta[] columns)
        {
            EnsureNotDisposed();
            if (columns is null) throw new ArgumentNullException(nameof(columns));
            EnsureSheetCapacity(_currentSheetIndex);
            var st = _sheets[_currentSheetIndex];
            st.Columns = columns;
            st.StyleIds = BuildStylePool(columns);
        }

        private const int MaxSheetNameLength = 31;
        private const int MaxRows = 1_048_576;
        private const int MaxColumns = 16_384;
        private static readonly char[] InvalidSheetNameChars = { ':', '\\', '/', '?', '*', '[', ']' };
        private const int SheetStreamBufferBytes = 64 * 1024;
        private const int SheetStreamFlushThresholdBytes = 16 * 1024;

        /// <summary>
        /// Creates a new worksheet with the given name and makes it the current sheet.
        /// </summary>
        public void AddSheet(string name)
        {
            EnsureNotDisposed();
            try
            {
                if (name is null) throw new ArgumentNullException(nameof(name));
                name = name.Trim();
                ValidateSheetName(name);
                if (_sheetNames.Any(existing => string.Equals(existing, name, StringComparison.OrdinalIgnoreCase)))
                    throw new ArgumentException($"A worksheet named '{name}' already exists.", nameof(name));

                if (_sheetStream is not null)
                    CloseCurrentSheet(_currentSheetIndex);

                _sheetNames.Add(name);
                _currentSheetIndex++;
                EnsureSheetCapacity(_currentSheetIndex);
                int sheetNo = _currentSheetIndex + 1;

                _sheetStream = _zip.OpenEntry($"xl/worksheets/sheet{sheetNo}.xml", _compression);

                _sink.WriteUtf8("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"u8);
                _sink.WriteUtf8("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\r\n"u8);
                _sink.FlushTo(_sheetStream);
            }
            catch (Exception ex)
            {
                RecordFault(ex);
                throw;
            }
        }

        /// <summary>
        /// Writes sheet-level metadata such as column widths, hidden columns, and the frozen header.
        /// </summary>
        public void WriteSheetMeta(ColumnMeta[] columns, bool freezeHeader)
        {
            EnsureNotDisposed();
            if (_sheetStream is null) throw new InvalidOperationException("Must call AddSheet before WriteSheetMeta");
            if (columns is null) throw new ArgumentNullException(nameof(columns));

            _sink.FlushTo(_sheetStream);

            WriteSheetPr(_sheets[_currentSheetIndex].Outline, _sheets[_currentSheetIndex].PageSetup);

            var sb = new System.Text.StringBuilder(256);

            if (freezeHeader)
            {
                sb.Append("<sheetViews><sheetView workbookViewId=\"0\"><pane ySplit=\"1\" topLeftCell=\"A2\" activePane=\"bottomLeft\" state=\"frozen\"/><selection pane=\"bottomLeft\"/></sheetView></sheetViews>\r\n");
            }

            if (_defaultRowHeight is double h2)
            {
                double pt = h2 * 0.75;
                sb.Append("<sheetFormatPr defaultRowHeight=\"").Append(pt.ToString("0.###", CultureInfo.InvariantCulture)).Append("\" customHeight=\"1\"/>\r\n");
            }

            bool hasColsMeta = false;
            for (int i = 0; i < columns.Length; i++)
            {
                if (columns[i].Width.HasValue || columns[i].Hidden) { hasColsMeta = true; break; }
            }
            if (hasColsMeta)
            {
                sb.Append("<cols>\r\n");
                int runStart = -1;
                double runWidth = 0;
                bool runHidden = false;
                int runMin = 0, runMax = 0;
                void FlushRun()
                {
                    if (runStart == -1) return;
                    if (runHidden)
                        sb.Append("<col min=\"").Append(runMin).Append("\" max=\"").Append(runMax).Append("\" hidden=\"1\"/>\r\n");
                    else
                        sb.Append("<col min=\"").Append(runMin).Append("\" max=\"").Append(runMax).Append("\" width=\"").Append(runWidth.ToString("R", CultureInfo.InvariantCulture)).Append("\" customWidth=\"1\"/>\r\n");
                }
                for (int i = 0; i < columns.Length; i++)
                {
                    var c = columns[i];
                    int colNo = i + 1;
                    bool isHidden = c.Hidden;
                    if (!isHidden && !c.Width.HasValue)
                    {
                        FlushRun();
                        runStart = -1;
                        continue;
                    }
                    if (runStart == -1)
                    {
                        runStart = colNo;
                        runWidth = c.Width ?? 0;
                        runHidden = isHidden;
                        runMin = colNo;
                        runMax = colNo;
                    }
                    else if (isHidden == runHidden && (isHidden || c.Width == runWidth))
                    {
                        runMax = colNo;
                    }
                    else
                    {
                        FlushRun();
                        runStart = colNo;
                        runWidth = c.Width ?? 0;
                        runHidden = isHidden;
                        runMin = colNo;
                        runMax = colNo;
                    }
                }
                FlushRun();
                sb.Append("</cols>\r\n");
            }

            if (sb.Length > 0)
            {
                _sink.WriteUtf8(sb);
            }

            _sink.FlushTo(_sheetStream);
        }

        /// <summary>
        /// Writes the header row from the supplied column metadata.
        /// </summary>
        public void WriteHeader(ColumnMeta[] columns)
        {
            EnsureNotDisposed();
            if (columns is null) throw new ArgumentNullException(nameof(columns));
            _currentColumns = columns;

            EnsureColumnLetters(columns.Length);

            if (!_sheetDataStarted)
            {
                _sink.WriteUtf8("<sheetData>\r\n"u8);
                _sheetDataStarted = true;
            }

            double? headerRowHeight = null;
            for (int i = 0; i < columns.Length; i++)
            {
                if (columns[i].RowHeight.HasValue) { headerRowHeight = columns[i].RowHeight; break; }
            }
            BeginRow(headerRowHeight);
            for (int i = 0; i < columns.Length; i++)
            {
                var c = columns[i];
                var colLetter = _colLetterCache![_currentColIndex];
                _currentColIndex++;
                _sink.WriteInlineStringCell(0, c.DisplayName, colLetter, RowNumberRef, _strictCellReferences);
            }
            EndRow();
        }

        /// <summary>
        /// Merges the given cell range in the current sheet.
        /// </summary>
        public void MergeCells(string range)
        {
            EnsureNotDisposed();
            if (range is null) throw new ArgumentNullException(nameof(range));
            if (_currentSheetIndex < 0) throw new InvalidOperationException("Must call AddSheet first");
            ValidateCellRange(range, nameof(range));
            EnsureSheetCapacity(_currentSheetIndex);
            _sheets[_currentSheetIndex].MergeCells.Add(range);
        }

        /// <summary>
        /// Sets the auto-filter range for the current sheet.
        /// </summary>
        public void SetAutoFilter(string cellRange)
        {
            EnsureNotDisposed();
            if (cellRange is null) throw new ArgumentNullException(nameof(cellRange));
            if (_currentSheetIndex < 0) throw new InvalidOperationException("Must call AddSheet before setting autoFilter");
            ValidateCellRange(cellRange, nameof(cellRange));
            EnsureSheetCapacity(_currentSheetIndex);
            _sheets[_currentSheetIndex].AutoFilter = cellRange;
        }

        /// <summary>
        /// Adds a hyperlink to a cell in the current sheet.
        /// </summary>
        public void AddHyperlink(string cellRef, string uri)
        {
            EnsureNotDisposed();
            if (cellRef is null) throw new ArgumentNullException(nameof(cellRef));
            if (uri is null) throw new ArgumentNullException(nameof(uri));
            if (CellRefToCol(cellRef) < 0) throw new ArgumentException($"invalid cellRef: '{cellRef}'", nameof(cellRef));
            if (CellRefToRow(cellRef) < 0) throw new ArgumentException($"invalid cellRef: '{cellRef}'", nameof(cellRef));
            if (!Uri.TryCreate(uri, UriKind.Absolute, out var parsed) ||
                (parsed.Scheme != "http" && parsed.Scheme != "https" && parsed.Scheme != "mailto" && parsed.Scheme != "file"))
                throw new ArgumentException($"URI must be an absolute http/https/mailto/file URI: '{uri}'", nameof(uri));
            if (_currentSheetIndex < 0) throw new InvalidOperationException("Must call AddSheet before adding hyperlink");
            EnsureSheetCapacity(_currentSheetIndex);
            _sheets[_currentSheetIndex].Hyperlinks.Add((cellRef, uri));
        }

        /// <summary>
        /// Adds a data-validation rule to the current sheet.
        /// </summary>
        public void AddDataValidation(DataValidation validation)
        {
            EnsureNotDisposed();
            if (validation is null) throw new ArgumentNullException(nameof(validation));
            if (_currentSheetIndex < 0) throw new InvalidOperationException("Must call AddSheet before adding data validation");
            ValidateCellRange(validation.CellRange, nameof(validation));
            EnsureSheetCapacity(_currentSheetIndex);
            _sheets[_currentSheetIndex].DataValidations.Add(validation);
        }

        /// <summary>
        /// Adds an Excel table definition to the current sheet.
        /// </summary>
        public void AddTable(TableDefinition table)
        {
            EnsureNotDisposed();
            if (table is null) throw new ArgumentNullException(nameof(table));
            if (_currentSheetIndex < 0) throw new InvalidOperationException("Must call AddSheet before adding table");
            ValidateTableDefinition(table);
            var refStr = table.Ref;
            int colon = refStr.IndexOf(':');
            if (colon < 0
                || CellRefToCol(refStr.Substring(0, colon)) < 0
                || CellRefToRow(refStr.Substring(0, colon)) < 0
                || CellRefToCol(refStr.Substring(colon + 1)) < 0
                || CellRefToRow(refStr.Substring(colon + 1)) < 0)
            {
                throw new ArgumentException($"Table Ref invalid (expected A1:B10 format): '{refStr}'", nameof(table));
            }
            EnsureSheetCapacity(_currentSheetIndex);
            _sheets[_currentSheetIndex].Tables.Add(table);
        }

        /// <summary>
        /// Applies protection options to the current sheet.
        /// </summary>
        public void SetSheetProtection(SheetProtection protection)
        {
            EnsureNotDisposed();
            if (protection is null) throw new ArgumentNullException(nameof(protection));
            if (_currentSheetIndex < 0) throw new InvalidOperationException("Must call AddSheet first");
            EnsureSheetCapacity(_currentSheetIndex);
            _sheets[_currentSheetIndex].Protection = protection;
        }

        /// <summary>
        /// Sets the print page setup for the current sheet.
        /// </summary>
        public void SetPageSetup(PageSetup ps)
        {
            EnsureNotDisposed();
            if (ps is null) throw new ArgumentNullException(nameof(ps));
            if (_currentSheetIndex < 0) throw new InvalidOperationException("Must call AddSheet first");
            EnsureSheetCapacity(_currentSheetIndex);
            _sheets[_currentSheetIndex].PageSetup = ps;
        }

        /// <summary>
        /// Sets the outline (grouping) options for the current sheet.
        /// </summary>
        public void SetOutline(OutlineSettings outline)
        {
            EnsureNotDisposed();
            if (outline is null) throw new ArgumentNullException(nameof(outline));
            if (_currentSheetIndex < 0) throw new InvalidOperationException("Must call AddSheet first");
            EnsureSheetCapacity(_currentSheetIndex);
            _sheets[_currentSheetIndex].Outline = outline;
        }

        /// <summary>
        /// Adds a cell comment to the current sheet.
        /// </summary>
        public void AddComment(Comment comment)
        {
            EnsureNotDisposed();
            if (comment is null) throw new ArgumentNullException(nameof(comment));
            if (_currentSheetIndex < 0) throw new InvalidOperationException("Must call AddSheet first");
            EnsureSheetCapacity(_currentSheetIndex);
            _sheets[_currentSheetIndex].Comments.Add(comment);
        }

        /// <summary>
        /// Adds a workbook- or sheet-scoped named range.
        /// </summary>
        public void AddNamedRange(NamedRange range)
        {
            EnsureNotDisposed();
            if (range is null) throw new ArgumentNullException(nameof(range));
            _namedRanges.Add(range);
        }

        /// <summary>
        /// Adds conditional formatting to the current sheet.
        /// </summary>
        public void AddConditionalFormatting(ConditionalFormatting cf)
        {
            EnsureNotDisposed();
            if (cf is null) throw new ArgumentNullException(nameof(cf));
            if (_currentSheetIndex < 0) throw new InvalidOperationException("Must call AddSheet first");
            ValidateCellRange(cf.CellRange, nameof(cf));
            if (cf.Rules.Length == 0) throw new ArgumentException("At least one conditional formatting rule is required.", nameof(cf));
            EnsureSheetCapacity(_currentSheetIndex);
            _sheets[_currentSheetIndex].ConditionalFormattings.Add(cf);
        }

        /// <summary>
        /// Adds an image to the current sheet and returns its one-based picture number.
        /// </summary>
        public int AddImage(byte[] imageBytes, string extension, string anchorFromCell, string anchorToCell)
        {
            EnsureNotDisposed();
            if (imageBytes is null || imageBytes.Length == 0) throw new ArgumentException("imageBytes empty", nameof(imageBytes));
            if (extension is null) throw new ArgumentNullException(nameof(extension));
            if (string.IsNullOrEmpty(extension)) extension = "png";
#if NETSTANDARD2_0
            if (extension.StartsWith(".")) extension = extension.Substring(1);
#else
            if (extension.StartsWith('.')) extension = extension[1..];
#endif
            if (extension != "png" && extension != "jpeg" && extension != "jpg" && extension != "gif" && extension != "bmp")
                throw new ArgumentException($"unsupported image extension '{extension}'(supported: png, jpeg, jpg, gif, bmp)", nameof(extension));
            if (_currentSheetIndex < 0) throw new InvalidOperationException("Must call AddSheet before AddImage");

            if (CellRefToCol(anchorFromCell) < 0) throw new ArgumentException($"invalid cellRef: '{anchorFromCell}'", nameof(anchorFromCell));
            if (CellRefToRow(anchorFromCell) < 0) throw new ArgumentException($"invalid cellRef: '{anchorFromCell}'", nameof(anchorFromCell));
            if (CellRefToCol(anchorToCell) < 0) throw new ArgumentException($"invalid cellRef: '{anchorToCell}'", nameof(anchorToCell));
            if (CellRefToRow(anchorToCell) < 0) throw new ArgumentException($"invalid cellRef: '{anchorToCell}'", nameof(anchorToCell));

            _imageCounter++;
            int imageNo = _imageCounter;
            string imageName = $"image{imageNo}.{extension}";

            _imageBytes.Add(($"xl/media/{imageName}", imageBytes));

            EnsureSheetCapacity(_currentSheetIndex);
            _sheets[_currentSheetIndex].Images.Add(new ImageAnchor(imageNo, imageName, anchorFromCell, anchorToCell));
            return imageNo;
        }

        /// <summary>
        /// Writes the data rows using a reflection-based row plan.
        /// </summary>
        public void WriteRows<T>(IEnumerable<T> data, RowPlan plan)
        {
            EnsureNotDisposed();
            if (data is null) throw new ArgumentNullException(nameof(data));
            if (plan is null) throw new ArgumentNullException(nameof(plan));

            try
            {
                int colCount = plan.Columns.Length;
                var getters = plan.Getters;
                var styleIds = GetCurrentSheetStyleIds(colCount);
                var columns = plan.Columns;
                EnsureColumnLetters(colCount);

                foreach (var item in data)
                {
                    if (item is null) continue;
                    BeginRow();
                    for (int i = 0; i < colCount; i++)
                    {
                        var cv = getters[i](item);
                        WriteCell(cv, styleIds[i]);
                    }
                    EndRow();
                }
            }
            catch (Exception ex)
            {
                RecordFault(ex);
                throw;
            }
        }

        /// <summary>
        /// Writes the data rows using a strongly typed row plan.
        /// </summary>
        public void WriteRows<T>(IEnumerable<T> data, TypedRowPlan<T> plan)
        {
            EnsureNotDisposed();
            if (data is null) throw new ArgumentNullException(nameof(data));
            if (plan is null) throw new ArgumentNullException(nameof(plan));

            try
            {
                int colCount = plan.Columns.Length;
                var getters = plan.TypedGetters;
                var cellWriters = plan.TypedCellWriters;
                var columns = plan.Columns;
                var styleIds = GetCurrentSheetStyleIds(colCount);
                EnsureColumnLetters(colCount);
                var formulas = new string?[colCount];
                for (int i = 0; i < colCount; i++) formulas[i] = columns[i].Formula;

                bool fastPath = cellWriters is not null && cellWriters.Length == colCount;
                // Generated writers avoid CellValue boxing. Formulas and unsupported columns use
                // the general path below.
                if (fastPath)
                {
                    for (int i = 0; i < colCount; i++)
                    {
                        if (formulas![i] is not null || cellWriters![i] is null)
                        {
                            fastPath = false;
                            break;
                        }
                    }
                }

                if (fastPath)
                {
                    foreach (var item in data)
                    {
                        if (item is null) continue;
                        var row = BeginRow();
                        for (int i = 0; i < colCount; i++)
                        {
                            cellWriters![i]!(row, item, styleIds[i]);
                        }
                        row.EndRow();
                    }
                }
                else
                {
                    foreach (var item in data)
                    {
                        if (item is null) continue;
                        var row = BeginRow();
                        for (int i = 0; i < colCount; i++)
                        {
                            if (formulas[i] is { } tpl)
                            {
                                row.WriteFormula(styleIds[i], tpl);
                            }
                            else if (cellWriters is not null && cellWriters[i] is { } cw)
                            {
                                cw(row, item, styleIds[i]);
                            }
                            else
                            {
                                var cv = getters[i](item);
                                row.WriteCell(cv, styleIds[i]);
                            }
                        }
                        row.EndRow();
                    }
                }
            }
            catch (Exception ex)
            {
                RecordFault(ex);
                throw;
            }
        }

        /// <summary>
        /// Writes data rows from an asynchronous enumeration using a strongly typed row plan.
        /// </summary>
        public async Task WriteRowsAsync<T>(IAsyncEnumerable<T> data, TypedRowPlan<T> plan, CancellationToken cancellationToken = default)
        {
            EnsureNotDisposed();
            if (data is null) throw new ArgumentNullException(nameof(data));
            if (plan is null) throw new ArgumentNullException(nameof(plan));

            try
            {
            int colCount = plan.Columns.Length;
            var getters = plan.TypedGetters;
            var columns = plan.Columns;
            var styleIds = GetCurrentSheetStyleIds(colCount);
            EnsureColumnLetters(colCount);
            var cellWriters = plan.TypedCellWriters;
            var formulas = new string?[colCount];
            for (int i = 0; i < colCount; i++) formulas[i] = columns[i].Formula;

            bool fastPath = cellWriters is not null && cellWriters.Length == colCount;
            // Keep the generated path allocation-free when every column has a direct writer.
            if (fastPath)
            {
                for (int i = 0; i < colCount; i++)
                {
                    if (formulas![i] is not null || cellWriters![i] is null)
                    {
                        fastPath = false;
                        break;
                    }
                }
            }

            if (fastPath)
            {
                var enumerator = data.GetAsyncEnumerator(cancellationToken);
                try
                {
                    while (true)
                    {
                        var moveNext = enumerator.MoveNextAsync();
                        bool hasMore = moveNext.IsCompletedSuccessfully
                            ? moveNext.Result
                            : await moveNext.ConfigureAwait(false);
                        if (!hasMore) break;
                        cancellationToken.ThrowIfCancellationRequested();
                        var item = enumerator.Current;
                        if (item is null) continue;
                        var row = BeginRow();
                        for (int i = 0; i < colCount; i++)
                        {
                            cellWriters![i]!(row, item, styleIds[i]);
                        }
                        if (EndRowCheckFlush())
                            await _sink.FlushToAsync(_sheetStream!, cancellationToken).ConfigureAwait(false);
                    }
                }
                finally
                {
                    await enumerator.DisposeAsync().ConfigureAwait(false);
                }
            }
            else
            {
                var enumerator = data.GetAsyncEnumerator(cancellationToken);
                try
                {
                    while (true)
                    {
                        var moveNext = enumerator.MoveNextAsync();
                        bool hasMore = moveNext.IsCompletedSuccessfully
                            ? moveNext.Result
                            : await moveNext.ConfigureAwait(false);
                        if (!hasMore) break;
                        cancellationToken.ThrowIfCancellationRequested();
                        var item = enumerator.Current;
                        if (item is null) continue;
                        var row = BeginRow();
                        for (int i = 0; i < colCount; i++)
                        {
                            if (formulas[i] is { } tpl)
                            {
                                row.WriteFormula(styleIds[i], tpl);
                            }
                            else if (cellWriters is not null && cellWriters[i] is { } cw)
                            {
                                cw(row, item, styleIds[i]);
                            }
                            else
                            {
                                var cv = getters[i](item);
                                row.WriteCell(cv, styleIds[i]);
                            }
                        }
                        if (EndRowCheckFlush())
                            await _sink.FlushToAsync(_sheetStream!, cancellationToken).ConfigureAwait(false);
                    }
                }
                finally
                {
                    await enumerator.DisposeAsync().ConfigureAwait(false);
                }
            }

            if (_sheetStream is not null)
            {
                await _sink.FlushToAsync(_sheetStream, cancellationToken).ConfigureAwait(false);
            }
            }
            catch (Exception ex)
            {
                RecordFault(ex);
                throw;
            }
        }

        /// <summary>
        /// Writes data rows from a materialized list using a strongly typed row plan, avoiding extra enumerator allocation.
        /// </summary>
        public async Task WriteRowsFromListAsync<T>(IList<T> data, TypedRowPlan<T> plan, CancellationToken cancellationToken = default)
        {
            EnsureNotDisposed();
            if (data is null) throw new ArgumentNullException(nameof(data));
            if (plan is null) throw new ArgumentNullException(nameof(plan));

            try
            {
            int colCount = plan.Columns.Length;
            var getters = plan.TypedGetters;
            var columns = plan.Columns;
            var styleIds = GetCurrentSheetStyleIds(colCount);
            EnsureColumnLetters(colCount);
            var cellWriters = plan.TypedCellWriters;
            var formulas = new string?[colCount];
            for (int i = 0; i < colCount; i++) formulas[i] = columns[i].Formula;

            bool fastPath = cellWriters is not null && cellWriters.Length == colCount;
            if (fastPath)
            {
                for (int i = 0; i < colCount; i++)
                {
                    if (formulas![i] is not null || cellWriters![i] is null)
                    {
                        fastPath = false;
                        break;
                    }
                }
            }

            int count = data.Count;
            for (int idx = 0; idx < count; idx++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var item = data[idx];
                if (item is null) continue;
                var row = BeginRow();
                if (fastPath)
                {
                    for (int i = 0; i < colCount; i++)
                        cellWriters![i]!(row, item, styleIds[i]);
                }
                else
                {
                    for (int i = 0; i < colCount; i++)
                    {
                        if (formulas[i] is { } tpl)
                            row.WriteFormula(styleIds[i], tpl);
                        else if (cellWriters is not null && cellWriters[i] is { } cw)
                            cw(row, item, styleIds[i]);
                        else
                            row.WriteCell(getters[i](item), styleIds[i]);
                    }
                }
                if (EndRowCheckFlush())
                    await _sink.FlushToAsync(_sheetStream!, cancellationToken).ConfigureAwait(false);
            }

            if (_sheetStream is not null)
                await _sink.FlushToAsync(_sheetStream, cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                RecordFault(ex);
                throw;
            }
        }

        private void WriteFormulaCell(int styleId, string template, int row)
        {
            _currentColIndex++;
            _sink.WriteUtf8("<c"u8);
            if (styleId > 0)
            {
                _sink.WriteUtf8(" s=\""u8);
                _sink.WriteInt32(styleId);
                _sink.WriteUtf8("\""u8);
            }
            _sink.WriteUtf8(" r=\""u8);
            var colLetterF = _colLetterCache![_currentColIndex - 1];
            _sink.WriteUtf8(colLetterF);
            _sink.WriteUtf8(RowNumberRef);
            _sink.WriteUtf8("\">"u8);
            _sink.WriteUtf8("<f>"u8);

            Span<byte> rowBuf = stackalloc byte[12];
            NumberFormatHelper.WriteInt32(row, rowBuf, out var rw); int rowLen = rw;

            int cursor = 0;
            const string Token = "{row}";
            while (true)
            {
                int idx = template.IndexOf(Token, cursor, StringComparison.Ordinal);
                if (idx < 0)
                {
                    if (cursor < template.Length) _sink.WriteEscaped(template.Substring(cursor));
                    break;
                }
                if (idx > cursor) _sink.WriteEscaped(template.Substring(cursor, idx - cursor));
                if (rowLen > 0) _sink.WriteUtf8(rowBuf.Slice(0, rowLen));
                cursor = idx + Token.Length;
            }
            _sink.WriteUtf8("</f></c>"u8);
        }

        /// <summary>
        /// Begins a new row. You must call <c>EndRow</c> on the returned writer when finished.
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never)]
        public XlsxRowWriter BeginRow(double? height = null)
        {
            if (_rowIndex >= MaxRows)
                throw new InvalidOperationException($"A worksheet cannot contain more than {MaxRows} rows.");
            _rowIndex++;
            _currentColIndex = 0;
            NumberFormatHelper.WriteInt32(_rowIndex, _rowNumberBuf, out _rowNumberLen);
            var rowRef = RowNumberRef;
            _sink.WriteUtf8("<row r=\""u8);
            _sink.WriteUtf8(rowRef);
            _sink.WriteUtf8("\""u8);
            var h = height ?? _nextRowHeight;
            _nextRowHeight = null;
            if (h.HasValue && (h.Value <= 0 || double.IsNaN(h.Value) || double.IsInfinity(h.Value))) h = null;
            if (h.HasValue)
            {
                double pt = h.Value * 0.75;
                _sink.WriteUtf8(" ht=\""u8);
                Span<byte> buf = stackalloc byte[32];
                NumberFormatHelper.WriteDouble(pt, 'G', buf, out var w2);
                _sink.WriteUtf8(buf.Slice(0, w2));
                _sink.WriteUtf8("\" customHeight=\"1\""u8);
            }
            _sink.WriteUtf8(">"u8);
            return new XlsxRowWriter(this);
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        private bool EndRowCheckFlush()
        {
            _sink.WriteUtf8("</row>"u8);
            return _sink.RemainingCapacity < SheetStreamFlushThresholdBytes && _sheetStream is not null;
        }

        private void EndRow()
        {
            if (EndRowCheckFlush())
            {
                _sink.FlushTo(_sheetStream!);
            }
        }

        private async Task EndRowAsync(CancellationToken cancellationToken)
        {
            if (EndRowCheckFlush())
            {
                await _sink.FlushToAsync(_sheetStream!, cancellationToken).ConfigureAwait(false);
            }
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        public readonly struct XlsxRowWriter
        {
            private readonly XlsxWriter _w;
            internal XlsxRowWriter(XlsxWriter w) => _w = w;

            /// <summary>
            /// Writes a string cell.
            /// </summary>
            public void WriteStringCell(int styleId, string? s) => _w.WriteStringCell(styleId, s);

            /// <summary>
            /// Writes a numeric cell.
            /// </summary>
            public void WriteNumberCell(int styleId, double d) => _w.WriteNumberCell(styleId, d);

            /// <summary>
            /// Writes a Boolean cell.
            /// </summary>
            public void WriteBoolCell(int styleId, bool b) => _w.WriteBoolCell(styleId, b);

            /// <summary>
            /// Writes an empty cell.
            /// </summary>
            public void WriteEmptyCell(int styleId) => _w.WriteEmptyCell(styleId);

            /// <summary>
            /// Writes a cell using the kind and value of the supplied <see cref="CellValue"/>.
            /// </summary>
            public void WriteCell(CellValue cv, int styleId) => _w.WriteCell(cv, styleId);

            /// <summary>
            /// Writes a formula that substitutes the current row number for <c>{row}</c>.
            /// </summary>
            public void WriteFormula(int styleId, string template) => _w.WriteFormulaCell(styleId, template, _w._rowIndex);

            /// <summary>
            /// Ends the current row.
            /// </summary>
            public void EndRow() => _w.EndRow();

            /// <summary>
            /// Asynchronously ends the current row.
            /// </summary>
            public Task EndRowAsync(CancellationToken cancellationToken = default) => _w.EndRowAsync(cancellationToken);
        }

        private void BeginCell(int styleId)
        {
            int colIdx = _currentColIndex;
            _currentColIndex++;
            Span<byte> buf = stackalloc byte[32];
            int p = 0;
            buf[p++] = (byte)'<';
            buf[p++] = (byte)'c';
            buf[p++] = (byte)' ';
            buf[p++] = (byte)'r';
            buf[p++] = (byte)'=';
            buf[p++] = (byte)'"';
            var colLetter = _colLetterCache![colIdx];
            colLetter.CopyTo(buf.Slice(p));
            p += colLetter.Length;
            RowNumberRef.CopyTo(buf.Slice(p)); p += _rowNumberLen;
            buf[p++] = (byte)'"';
            if (styleId > 0)
            {
                buf[p++] = (byte)' ';
                buf[p++] = (byte)'s';
                buf[p++] = (byte)'=';
                buf[p++] = (byte)'"';
                NumberFormatHelper.WriteInt32(styleId, buf.Slice(p), out int w); p += w;
                buf[p++] = (byte)'"';
            }
            buf[p++] = (byte)'>';
            _sink.WriteUtf8(buf.Slice(0, p));
        }

        private void EndCell() => _sink.WriteUtf8("</c>"u8);

        private int ResolveSharedStringIndex(string s)
        {
            _sstRefCount++;
            if (_sharedStringLookup.TryGetValue(s, out var idx)) return idx;
            idx = _sharedStrings.Count;
            _sharedStrings.Add(s);
            _sharedStringLookup[s] = idx;
            return idx;
        }

        private void WriteSstCell(int styleId, int sstIndex)
        {
            int colIdx = _currentColIndex;
            _currentColIndex++;
            Span<byte> buf = stackalloc byte[48];
            int p = 0;
            buf[p++] = (byte)'<';
            buf[p++] = (byte)'c';
            buf[p++] = (byte)' ';
            buf[p++] = (byte)'t';
            buf[p++] = (byte)'=';
            buf[p++] = (byte)'"';
            buf[p++] = (byte)'s';
            buf[p++] = (byte)'"';
            buf[p++] = (byte)' ';
            buf[p++] = (byte)'r';
            buf[p++] = (byte)'=';
            buf[p++] = (byte)'"';
            var colLetterSst = _colLetterCache![colIdx];
            colLetterSst.CopyTo(buf.Slice(p));
            p += colLetterSst.Length;
            RowNumberRef.CopyTo(buf.Slice(p)); p += _rowNumberLen;
            buf[p++] = (byte)'"';
            if (styleId > 0)
            {
                buf[p++] = (byte)' ';
                buf[p++] = (byte)'s';
                buf[p++] = (byte)'=';
                buf[p++] = (byte)'"';
                NumberFormatHelper.WriteInt32(styleId, buf.Slice(p), out int w); p += w;
                buf[p++] = (byte)'"';
            }
            buf[p++] = (byte)'>';
            buf[p++] = (byte)'<';
            buf[p++] = (byte)'v';
            buf[p++] = (byte)'>';
            NumberFormatHelper.WriteInt32(sstIndex, buf.Slice(p), out int w2); p += w2;
            buf[p++] = (byte)'<';
            buf[p++] = (byte)'/';
            buf[p++] = (byte)'v';
            buf[p++] = (byte)'>';
            buf[p++] = (byte)'<';
            buf[p++] = (byte)'/';
            buf[p++] = (byte)'c';
            buf[p++] = (byte)'>';
            _sink.WriteUtf8(buf.Slice(0, p));
        }

        private void WriteCell(CellValue cv, int styleId)
        {
            if (_useSharedStrings && cv.Type == CellType.String && cv.StringValue is not null)
            {
                int sstIdx = ResolveSharedStringIndex(cv.StringValue);
                WriteSstCell(styleId, sstIdx);
                return;
            }
            if (cv.Type == CellType.String && cv.StringValue is not null)
            {
                var colLetter = _colLetterCache![_currentColIndex];
                _currentColIndex++;
                _sink.WriteInlineStringCell(styleId, cv.StringValue, colLetter, RowNumberRef, _strictCellReferences);
                return;
            }
            if (cv.Type == CellType.Number)
            {
                if (double.IsNaN(cv.NumberValue) || double.IsInfinity(cv.NumberValue))
                {
                    BeginCell(styleId);
                    EndCell();
                    return;
                }
                BeginNumberCell(styleId);
                WriteNumber(cv.NumberValue);
                EndCell();
                return;
            }
            if (cv.Type == CellType.Boolean)
            {
                WriteBoolCell(styleId, cv.BoolValue);
                return;
            }
            BeginCell(styleId);
            WriteCellValue(cv);
            EndCell();
        }

        private void WriteCellValue(CellValue cv)
        {
            switch (cv.Type)
            {
                case CellType.Null:
                    break;
                case CellType.String:
                    if (cv.StringValue is not null) WriteInlineString(cv.StringValue);
                    break;
                case CellType.Number:
                    WriteNumber(cv.NumberValue);
                    break;
                case CellType.Boolean:
                    _sink.WriteUtf8(cv.BoolValue ? "<v>1</v>"u8 : "<v>0</v>"u8);
                    break;
                case CellType.Formula:
                    _sink.WriteUtf8("<f>"u8);
                    _sink.WriteEscaped(cv.StringValue ?? "");
                    _sink.WriteUtf8("</f>"u8);
                    break;
            }
        }

        private void WriteInlineString(string s)
        {
            _sink.WriteUtf8("<is><t"u8);
            bool preserve = s.Length > 0 && (char.IsWhiteSpace(s[0]) || char.IsWhiteSpace(s[s.Length - 1]));
            if (preserve) _sink.WriteUtf8(" xml:space=\"preserve\""u8);
            _sink.WriteUtf8(">"u8);
            _sink.WriteEscaped(s);
            _sink.WriteUtf8("</t></is>"u8);
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        private void WriteStringCell(int styleId, string? s)
        {
            if (s is null)
            {
                BeginCell(styleId);
                EndCell();
                return;
            }
            if (_useSharedStrings)
            {
                int sstIdx = ResolveSharedStringIndex(s);
                WriteSstCell(styleId, sstIdx);
                return;
            }
            var colLetter = _colLetterCache![_currentColIndex];
            _currentColIndex++;
            _sink.WriteInlineStringCell(styleId, s, colLetter, RowNumberRef, _strictCellReferences);
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        private void WriteNumberCell(int styleId, double d)
        {
            if (double.IsNaN(d) || double.IsInfinity(d))
            {
                BeginCell(styleId);
                EndCell();
                return;
            }
            int colIdx = _currentColIndex;
            _currentColIndex++;
            var colLetter = _colLetterCache![colIdx];
            _sink.WriteNumberCell(styleId, colLetter, RowNumberRef, d, _strictCellReferences);
        }

        private void BeginNumberCell(int styleId)
        {
            int colIdx = _currentColIndex;
            _currentColIndex++;
            Span<byte> buf = stackalloc byte[40];
            int p = 0;
            buf[p++] = (byte)'<';
            buf[p++] = (byte)'c';
            buf[p++] = (byte)' ';
            buf[p++] = (byte)'r';
            buf[p++] = (byte)'=';
            buf[p++] = (byte)'"';
            var colLetterNum = _colLetterCache![colIdx];
            colLetterNum.CopyTo(buf.Slice(p));
            p += colLetterNum.Length;
            RowNumberRef.CopyTo(buf.Slice(p)); p += _rowNumberLen;
            buf[p++] = (byte)'"';
            buf[p++] = (byte)' ';
            buf[p++] = (byte)'t';
            buf[p++] = (byte)'=';
            buf[p++] = (byte)'"';
            buf[p++] = (byte)'n';
            buf[p++] = (byte)'"';
            if (styleId > 0)
            {
                buf[p++] = (byte)' ';
                buf[p++] = (byte)'s';
                buf[p++] = (byte)'=';
                buf[p++] = (byte)'"';
                NumberFormatHelper.WriteInt32(styleId, buf.Slice(p), out int w); p += w;
                buf[p++] = (byte)'"';
            }
            buf[p++] = (byte)'>';
            _sink.WriteUtf8(buf.Slice(0, p));
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        private void WriteBoolCell(int styleId, bool b)
        {
            int colIdx = _currentColIndex;
            _currentColIndex++;
            Span<byte> buf = stackalloc byte[48];
            int p = 0;
            buf[p++] = (byte)'<';
            buf[p++] = (byte)'c';
            buf[p++] = (byte)' ';
            buf[p++] = (byte)'t';
            buf[p++] = (byte)'=';
            buf[p++] = (byte)'"';
            buf[p++] = (byte)'b';
            buf[p++] = (byte)'"';
            buf[p++] = (byte)' ';
            buf[p++] = (byte)'r';
            buf[p++] = (byte)'=';
            buf[p++] = (byte)'"';
            var colLetterBool = _colLetterCache![colIdx];
            colLetterBool.CopyTo(buf.Slice(p));
            p += colLetterBool.Length;
            RowNumberRef.CopyTo(buf.Slice(p)); p += _rowNumberLen;
            buf[p++] = (byte)'"';
            if (styleId > 0)
            {
                buf[p++] = (byte)' ';
                buf[p++] = (byte)'s';
                buf[p++] = (byte)'=';
                buf[p++] = (byte)'"';
                NumberFormatHelper.WriteInt32(styleId, buf.Slice(p), out int w); p += w;
                buf[p++] = (byte)'"';
            }
            buf[p++] = (byte)'>';
            buf[p++] = (byte)'<';
            buf[p++] = (byte)'v';
            buf[p++] = (byte)'>';
            buf[p++] = b ? (byte)'1' : (byte)'0';
            buf[p++] = (byte)'<';
            buf[p++] = (byte)'/';
            buf[p++] = (byte)'v';
            buf[p++] = (byte)'>';
            buf[p++] = (byte)'<';
            buf[p++] = (byte)'/';
            buf[p++] = (byte)'c';
            buf[p++] = (byte)'>';
            _sink.WriteUtf8(buf.Slice(0, p));
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        private void WriteEmptyCell(int styleId)
        {
            BeginCell(styleId);
            EndCell();
        }

        private void WriteNumber(double d)
        {
            if (double.IsNaN(d) || double.IsInfinity(d)) return;
            _sink.WriteUtf8("<v>"u8);
            _sink.WriteDouble(d);
            _sink.WriteUtf8("</v>"u8);
        }

        private void CloseCurrentSheetCore(int closingSheetIdx)
        {
            if (_sheetStream is null) return;
            var st = _sheets[closingSheetIdx];
            bool hasImages = st.Images.Count > 0;

            if (_sheetDataStarted)
            {
                _sink.WriteUtf8("</sheetData>\r\n"u8);
                _sheetDataStarted = false;
            }
            WriteSheetProtection(st.Protection);
            if (st.AutoFilter is { } afRef)
            {
                _sink.WriteUtf8("<autoFilter ref=\""u8);
                _sink.WriteEscaped(afRef);
                _sink.WriteUtf8("\"/>\r\n"u8);
            }
            if (st.MergeCells.Count > 0)
            {
                var cells = st.MergeCells;
                _sink.WriteUtf8("<mergeCells count=\""u8);
                _sink.WriteInt32(cells.Count);
                _sink.WriteUtf8("\">\r\n"u8);
                foreach (var range in cells)
                {
                    _sink.WriteUtf8("<mergeCell ref=\""u8);
                    _sink.WriteEscaped(range);
                    _sink.WriteUtf8("\"/>\r\n"u8);
                }
                _sink.WriteUtf8("</mergeCells>\r\n"u8);
            }
            WriteConditionalFormatting(st.ConditionalFormattings);
            if (st.DataValidations.Count > 0)
                WriteDataValidationsToSink(st.DataValidations);
            if (st.Hyperlinks.Count > 0)
            {
                var hlinks = st.Hyperlinks;
                _sink.WriteUtf8("<hyperlinks>\r\n"u8);
                for (int hi = 0; hi < hlinks.Count; hi++)
                {
                    int rid = hi + 1;
                    _sink.WriteUtf8("<hyperlink ref=\""u8);
                    _sink.WriteEscaped(hlinks[hi].Ref);
                    _sink.WriteUtf8("\" r:id=\"rIdH"u8);
                    _sink.WriteInt32(rid);
                    _sink.WriteUtf8("\"/>\r\n"u8);
                }
                _sink.WriteUtf8("</hyperlinks>\r\n"u8);
            }
            WritePageMargins(st.PageSetup);
            WritePageSetup(st.PageSetup);
            WriteHeaderFooter(st.PageSetup);
            if (hasImages)
            {
                _sink.WriteUtf8("<drawing r:id=\"rIdImage"u8);
                _sink.WriteInt32(closingSheetIdx + 1);
                _sink.WriteUtf8("\"/>\r\n"u8);
            }
            if (st.Comments.Count > 0)
            {
                _sink.WriteUtf8("<legacyDrawing r:id=\"rIdVml"u8);
                _sink.WriteInt32(closingSheetIdx + 1);
                _sink.WriteUtf8("\"/>\r\n"u8);
            }
            if (st.Tables.Count > 0)
                WriteTablePartsToSink(st.Tables, closingSheetIdx + 1);

            _sink.WriteUtf8("</worksheet>"u8);
        }

        private void CloseCurrentSheet(int closingSheetIdx)
        {
            if (_sheetStream is null) return;
            CloseCurrentSheetCore(closingSheetIdx);
            _sink.FlushTo(_sheetStream);
            _sheetStream.Dispose();
            _sheetStream = null;
            _rowIndex = 0;
        }

        private async Task CloseCurrentSheetAsync(int closingSheetIdx, CancellationToken cancellationToken)
        {
            if (_sheetStream is null) return;
            CloseCurrentSheetCore(closingSheetIdx);
            await _sink.FlushToAsync(_sheetStream, cancellationToken).ConfigureAwait(false);
            await ForwardOnlyZipWriter.DisposeEntryStreamAsync(_sheetStream).ConfigureAwait(false);
            _sheetStream = null;
            _rowIndex = 0;
        }

        private int[] GetCurrentSheetStyleIds(int colCount)
        {
            if (_currentSheetIndex >= 0 && _sheets.Count > _currentSheetIndex && _sheets[_currentSheetIndex].StyleIds is { } ids)
                return ids;
            return new int[colCount];
        }

        private void EnsureColumnLetters(int colCount)
        {
            if (colCount > MaxColumns)
                throw new ArgumentOutOfRangeException(nameof(colCount), $"A worksheet cannot contain more than {MaxColumns} columns.");
            if (_colLetterCache is not null && _colLetterCache.Length >= colCount) return;
            var cache = new byte[colCount][];
            Span<byte> tmp = stackalloc byte[4];
            for (int c = 0; c < colCount; c++)
            {
                int len = ColumnLetter(c, tmp);
                var arr = new byte[len];
                tmp.Slice(0, len).CopyTo(arr);
                cache[c] = arr;
            }
            _colLetterCache = cache;
        }

        internal static int ColumnLetter(int col0, Span<byte> dest)
        {
            int n = col0;
            int p = dest.Length;
            do
            {
                dest[--p] = (byte)('A' + n % 26);
                n = n / 26 - 1;
            } while (n >= 0);
            int written = dest.Length - p;
            dest.Slice(p, written).CopyTo(dest);
            return written;
        }

        private static string CfOpName(CfOperator op) => op switch
        {
            CfOperator.NotEqual => "notEqual",
            CfOperator.GreaterThan => "greaterThan",
            CfOperator.GreaterThanOrEqual => "greaterThanOrEqual",
            CfOperator.LessThan => "lessThan",
            CfOperator.LessThanOrEqual => "lessThanOrEqual",
            CfOperator.Between => "between",
            CfOperator.NotBetween => "notBetween",
            _ => "equal",
        };

        private void BuildSharedStringsXml(out StringBuilder sb)
        {
            sb = new StringBuilder(4096);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
            sb.Append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"")
              .Append(_sstRefCount).Append("\" uniqueCount=\"").Append(_sharedStrings.Count).Append("\">\r\n");
            for (int i = 0; i < _sharedStrings.Count; i++)
            {
                var s = _sharedStrings[i];
                sb.Append("  <si><t");
                bool preserve = s.Length > 0 && (char.IsWhiteSpace(s[0]) || char.IsWhiteSpace(s[s.Length - 1]));
                if (preserve) sb.Append(" xml:space=\"preserve\"");
                sb.Append(">").Append(XmlHelper.EscapeXmlAttr(s)).Append("</t></si>\r\n");
            }
            sb.Append("</sst>\r\n");
        }

        private void WriteSharedStringsTable()
        {
            BuildSharedStringsXml(out var sb);
            WriteEntryStringBuilder("xl/sharedStrings.xml", sb);
        }

        private async Task WriteSharedStringsTableAsync(CancellationToken cancellationToken)
        {
            BuildSharedStringsXml(out var sb);
            await WriteEntryStringBuilderAsync("xl/sharedStrings.xml", sb, cancellationToken).ConfigureAwait(false);
        }

        private static void ValidateColumnMeta(ColumnMeta column)
        {
            if (column.Width is double width && (double.IsNaN(width) || double.IsInfinity(width) || width <= 0))
                throw new ArgumentOutOfRangeException(nameof(column.Width), "Column width must be finite and greater than 0.");
            if (column.FontSize is float fontSize && (float.IsNaN(fontSize) || float.IsInfinity(fontSize) || fontSize <= 0))
                throw new ArgumentOutOfRangeException(nameof(column.FontSize), "Font size must be finite and greater than 0.");
            if (column.RowHeight is double rowHeight && (double.IsNaN(rowHeight) || double.IsInfinity(rowHeight) || rowHeight <= 0))
                throw new ArgumentOutOfRangeException(nameof(column.RowHeight), "Row height must be finite and greater than 0.");
            ValidateColor(column.BackgroundColor, nameof(column.BackgroundColor));
            ValidateColor(column.FontColor, nameof(column.FontColor));
            ValidateColor(column.BorderColor, nameof(column.BorderColor));
        }

        private static void ValidateColor(string? color, string parameterName)
        {
            if (color is null) return;
            if ((color.Length != 6 && color.Length != 8) || color.Any(c => !Uri.IsHexDigit(c)))
                throw new ArgumentException("Color must be a 6- or 8-digit hexadecimal value.", parameterName);
        }

        private int[] BuildStylePool(ColumnMeta[] columns)
        {
            var xfIds = new int[columns.Length];
            _stylePool ??= new List<string>();
            _stylePoolLookup ??= new Dictionary<string, int>();
            _fontSet ??= new();
            _fillSet ??= new();
            _borderSet ??= new();
            _fontLookup ??= new Dictionary<FontKey, int>();
            _fillLookup ??= new Dictionary<string, int>();
            _borderLookup ??= new Dictionary<string, int>();
            _numFmtLookup ??= new Dictionary<string, int>();

            if (_stylePool.Count == 0)
            {
                const string defaultXfKey = "0|0|0|0|0|0|0";
                _stylePool.Add(defaultXfKey);
                _stylePoolLookup[defaultXfKey] = 0;
            }
            if (_fontSet.Count == 0)
            {
                var defaultFont = (false, 11f, (string?)null, "Calibri", false, false, false);
                _fontSet.Add(defaultFont);
                _fontLookup[new FontKey(defaultFont.Item1, defaultFont.Item2, defaultFont.Item3,
                    defaultFont.Item4, defaultFont.Item5, defaultFont.Item6, defaultFont.Item7)] = 0;
            }
            // fills: index 0 = none, index 1 = gray125 (serialized specially in BuildFullStylePoolXml);
            // ensure both slots exist so background-color fills start at index 2.
            while (_fillSet.Count < 2) _fillSet.Add(null);

            if (_fillLookup.Count == 0)
            {
                _fillLookup[""] = 0;
            }

            for (int n = 0; n < _numFmts.Count; n++)
            {
                if (!_numFmtLookup!.ContainsKey(_numFmts[n]))
                    _numFmtLookup[_numFmts[n]] = n;
            }
            if (_borderSet.Count == 0) _borderSet.Add(null);

            for (int i = 0; i < columns.Length; i++)
            {
                var c = columns[i];

                var font = (Bold: c.Bold == true, Size: c.FontSize ?? 11f, Color: c.FontColor, Name: c.FontName ?? "Calibri",
                            Italic: c.Italic == true, Underline: c.Underline == true, Strike: c.StrikeThrough == true);
                var fontKey = new FontKey(font.Bold, font.Size, font.Color, font.Name, font.Italic, font.Underline, font.Strike);
                int fontId;
                if (_fontLookup.TryGetValue(fontKey, out var fid)) fontId = fid;
                else
                {
                    fontId = _fontSet.Count;
                    _fontSet.Add(font);
                    _fontLookup[fontKey] = fontId;
                }

                int fillId = 0;
                if (!string.IsNullOrEmpty(c.BackgroundColor))
                {
                    var fillKey = c.BackgroundColor!;
                    if (_fillLookup.TryGetValue(fillKey, out var fillid)) fillId = fillid;
                    else
                    {
                        fillId = _fillSet.Count;
                        _fillSet.Add(c.BackgroundColor);
                        _fillLookup[fillKey] = fillId;
                    }
                }

                int borderId = 0;
                if (c.BorderStyle.HasValue)
                {
                    var bk = $"{(int)c.BorderStyle.Value}|{c.BorderColor ?? ""}";
                    if (_borderLookup.TryGetValue(bk, out var bid)) borderId = bid;
                    else
                    {
                        borderId = _borderSet.Count;
                        _borderSet.Add(bk);
                        _borderLookup[bk] = borderId;
                    }
                }

                int numFmtId = 0;
                var format = c.Format;
                if (format is { Length: > 0 })
                {
                    if (_numFmtLookup!.TryGetValue(format, out var nid)) numFmtId = 164 + nid;
                }

                int wrap = c.Wrap == true ? 1 : 0;
                int autoCenter = c.AutoCenter == true ? 1 : 0;
                int vAlign = c.VerticalAlignment.HasValue ? (int)c.VerticalAlignment.Value : 0;
                var xk = $"{fontId}|{fillId}|{borderId}|{numFmtId}|{wrap}|{autoCenter}|{vAlign}";
                int xfId;
                if (_stylePoolLookup.TryGetValue(xk, out var xid)) xfId = xid;
                else
                {
                    xfId = _stylePool.Count;
                    _stylePool.Add(xk);
                    _stylePoolLookup[xk] = xfId;
                }
                xfIds[i] = xfId;
            }
            return xfIds;
        }

        private void WriteStyles()
        {
            EmitStylePoolXml();
        }

        private Task WriteStylesAsync(CancellationToken cancellationToken)
        {
            return EmitStylePoolXmlAsync(cancellationToken);
        }

        private List<string>? _stylePool;
        private Dictionary<string, int>? _stylePoolLookup;
        private List<(bool Bold, float Size, string? Color, string Name, bool Italic, bool Underline, bool Strike)>? _fontSet;
        private Dictionary<FontKey, int>? _fontLookup;
        private List<string?>? _fillSet;
        private Dictionary<string, int>? _fillLookup;
        private List<string?>? _borderSet;
        private Dictionary<string, int>? _borderLookup;
        private Dictionary<string, int>? _numFmtLookup;

        private readonly struct FontKey : IEquatable<FontKey>
        {
            public readonly bool Bold;
            public readonly float Size;
            public readonly string? Color;
            public readonly string Name;
            public readonly bool Italic;
            public readonly bool Underline;
            public readonly bool Strike;
            public FontKey(bool bold, float size, string? color, string name, bool italic, bool underline, bool strike)
            { Bold = bold; Size = size; Color = color; Name = name; Italic = italic; Underline = underline; Strike = strike; }
            public bool Equals(FontKey other) =>
                Bold == other.Bold && Size == other.Size && Color == other.Color && Name == other.Name &&
                Italic == other.Italic && Underline == other.Underline && Strike == other.Strike;
            public override bool Equals(object? obj) => obj is FontKey f && Equals(f);
            public override int GetHashCode() =>
#if NETSTANDARD2_0
                ((Bold ? 1 : 0) * 397) ^ Size.GetHashCode() ^ (Color?.GetHashCode() ?? 0)
                    ^ (Name?.GetHashCode() ?? 0) ^ ((Italic ? 1 : 0) * 17)
                    ^ ((Underline ? 1 : 0) * 23) ^ ((Strike ? 1 : 0) * 31);
#else
                HashCode.Combine(Bold, Size, Color, Name, Italic, Underline, Strike);
#endif
        }

        private static readonly string _defaultStylesXml =
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n" +
            "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\r\n" +
            "  <fonts count=\"1\">\r\n" +
            "    <font><sz val=\"11\"/><name val=\"Calibri\"/></font>\r\n" +
            "  </fonts>\r\n" +
            "  <fills count=\"2\">\r\n" +
            "    <fill><patternFill patternType=\"none\"/></fill>\r\n" +
            "    <fill><patternFill patternType=\"gray125\"/></fill>\r\n" +
            "  </fills>\r\n" +
            "  <borders count=\"1\">\r\n" +
            "    <border/>\r\n" +
            "  </borders>\r\n" +
            "  <cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>\r\n" +
            "  <cellXfs count=\"1\">\r\n" +
            "    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>\r\n" +
            "  </cellXfs>\r\n" +
            "  <cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>\r\n" +
            "</styleSheet>\r\n";

        private void WriteDefaultStylesToSink()
        {
            using var es = _zip.OpenEntry("xl/styles.xml", _compression);
            _sink.WriteUtf8(_defaultStylesXml);
            _sink.FlushTo(es);
        }

        private async Task WriteDefaultStylesToSinkAsync(CancellationToken cancellationToken)
        {
            var es = _zip.OpenEntry("xl/styles.xml", _compression);
            try
            {
                _sink.WriteUtf8(_defaultStylesXml);
                await _sink.FlushToAsync(es, cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                await ForwardOnlyZipWriter.DisposeEntryStreamAsync(es).ConfigureAwait(false);
            }
        }

        private void BuildFullStylePoolXml(out StringBuilder sb)
        {
            sb = new StringBuilder(2048);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
            sb.Append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\r\n");

            if (_numFmts.Count > 0)
            {
                sb.Append("  <numFmts count=\"").Append(_numFmts.Count).Append("\">\r\n");
                for (int i = 0; i < _numFmts.Count; i++)
                {
                    sb.Append("    <numFmt numFmtId=\"").Append(164 + i).Append("\" formatCode=\"")
                      .Append(XmlHelper.EscapeXmlAttr(_numFmts[i])).Append("\"/>\r\n");
                }
                sb.Append("  </numFmts>\r\n");
            }

            sb.Append("  <fonts count=\"").Append(_fontSet!.Count).Append("\">\r\n");
            foreach (var font in _fontSet)
            {
                var bold = font.Bold;
                var size = font.Size;
                var color = font.Color;
                var name = font.Name;
                var italic = font.Italic;
                var underline = font.Underline;
                var strike = font.Strike;
                sb.Append("    <font>");
                sb.Append("<sz val=\"").Append(size.ToString("0.#", CultureInfo.InvariantCulture)).Append("\"/>");
                sb.Append("<name val=\"").Append(XmlHelper.EscapeXmlAttr(name)).Append("\"/>");
                if (bold) sb.Append("<b/>");
                if (italic) sb.Append("<i/>");
                if (strike) sb.Append("<strike/>");
                if (color is not null)
                {
                    sb.Append("<color rgb=\"").Append(XmlHelper.EscapeXmlAttr(color)).Append("\"/>");
                }
                if (underline) sb.Append("<u val=\"single\"/>");
                sb.Append("</font>\r\n");
            }
            sb.Append("  </fonts>\r\n");

            sb.Append("  <fills count=\"").Append(_fillSet!.Count).Append("\">\r\n");
            for (int i = 0; i < _fillSet.Count; i++)
            {
                if (i == 1)
                {
                    sb.Append("    <fill><patternFill patternType=\"gray125\"/></fill>\r\n");
                    continue;
                }
                var color = _fillSet[i];
                if (color is null)
                {
                    sb.Append("    <fill><patternFill patternType=\"none\"/></fill>\r\n");
                }
                else
                {
                    sb.Append("    <fill><patternFill patternType=\"solid\"><fgColor rgb=\"").Append(XmlHelper.EscapeXmlAttr(color)).Append("\"/><bgColor rgb=\"").Append(XmlHelper.EscapeXmlAttr(color)).Append("\"/></patternFill></fill>\r\n");
                }
            }
            sb.Append("  </fills>\r\n");

            sb.Append("  <borders count=\"").Append(_borderSet!.Count).Append("\">\r\n");
            for (int i = 0; i < _borderSet.Count; i++)
            {
                var bk = _borderSet[i];
                if (i == 0 || bk is null)
                {
                    sb.Append("    <border/>\r\n");
                }
                else
                {
                    var parts = bk.Split('|');
                    var styleName = parts[0] switch
                    {
                        "1" => "thin",
                        "2" => "medium",
                        "3" => "thick",
                        "4" => "dashed",
                        "5" => "dotted",
                        "6" => "double",
                        _ => "thin",
                    };
                    string colorAttr = $"<color rgb=\"{XmlHelper.EscapeXmlAttr(parts[1])}\"/>";
                    sb.Append("    <border>");
                    sb.Append("<left style=\"").Append(styleName).Append("\">").Append(colorAttr).Append("</left>");
                    sb.Append("<right style=\"").Append(styleName).Append("\">").Append(colorAttr).Append("</right>");
                    sb.Append("<top style=\"").Append(styleName).Append("\">").Append(colorAttr).Append("</top>");
                    sb.Append("<bottom style=\"").Append(styleName).Append("\">").Append(colorAttr).Append("</bottom>");
                    sb.Append("</border>\r\n");
                }
            }
            sb.Append("  </borders>\r\n");

            sb.Append("  <cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>\r\n");

            sb.Append("  <cellXfs count=\"").Append(_stylePool!.Count).Append("\">\r\n");
            sb.Append("    <xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>\r\n");
            for (int i = 1; i < _stylePool.Count; i++)
            {
                var parts = _stylePool[i].Split('|');
                int fontId = int.Parse(parts[0]);
                int fillId = int.Parse(parts[1]);
                int borderId = int.Parse(parts[2]);
                int numFmtId = int.Parse(parts[3]);
                bool wrap = parts[4] == "1";
                bool autoCenter = parts[5] == "1";
                int vAlign = parts.Length > 6 ? int.Parse(parts[6]) : 0;
                string align = autoCenter ? "center" : "general";
                bool hasAlign = wrap || autoCenter || vAlign > 0;
                sb.Append("    <xf numFmtId=\"").Append(numFmtId).Append("\" fontId=\"").Append(fontId)
                  .Append("\" fillId=\"").Append(fillId).Append("\" borderId=\"").Append(borderId).Append("\" xfId=\"0\"");
                if (numFmtId > 0) sb.Append(" applyNumberFormat=\"1\"");
                if (fontId > 0) sb.Append(" applyFont=\"1\"");
                if (fillId > 0) sb.Append(" applyFill=\"1\"");
                if (borderId > 0) sb.Append(" applyBorder=\"1\"");
                if (hasAlign) sb.Append(" applyAlignment=\"1\"");
                if (hasAlign)
                {
                    sb.Append("><alignment");
                    if (wrap) sb.Append(" wrapText=\"1\"");
                    if (autoCenter) sb.Append(" horizontal=\"").Append(align).Append("\"");
                    if (vAlign > 0)
                    {
                        string va = vAlign switch { 2 => "center", 3 => "bottom", _ => "top" };
                        sb.Append(" vertical=\"").Append(va).Append("\"");
                    }
                    sb.Append("/></xf>\r\n");
                }
                else
                {
                    sb.Append("/>\r\n");
                }
            }
            sb.Append("  </cellXfs>\r\n");

            sb.Append("  <cellStyles count=\"1\"><cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>\r\n");
            sb.Append("</styleSheet>\r\n");
        }

        private void EmitStylePoolXml()
        {
            if (_stylePool is null || _fontSet is null || _fillSet is null || _borderSet is null)
            {
                WriteDefaultStylesToSink();
                return;
            }

#if !DEBUG
            if (_numFmts.Count == 0 && _stylePool.Count == 1)
            {
                WriteDefaultStylesToSink();
                return;
            }
#endif
            BuildFullStylePoolXml(out var sb);
            WriteEntryStringBuilder("xl/styles.xml", sb);

#if DEBUG
            if (_numFmts.Count == 0 && _stylePool.Count == 1 &&
                !string.Equals(sb.ToString(), _defaultStylesXml, StringComparison.Ordinal))
            {
                throw new InvalidOperationException("Default styles.xml template is out of sync with the builder.");
            }
#endif
        }

        private async Task EmitStylePoolXmlAsync(CancellationToken cancellationToken)
        {
            if (_stylePool is null || _fontSet is null || _fillSet is null || _borderSet is null)
            {
                await WriteDefaultStylesToSinkAsync(cancellationToken).ConfigureAwait(false);
                return;
            }

#if !DEBUG
            if (_numFmts.Count == 0 && _stylePool.Count == 1)
            {
                await WriteDefaultStylesToSinkAsync(cancellationToken).ConfigureAwait(false);
                return;
            }
#endif
            BuildFullStylePoolXml(out var sb);
            await WriteEntryStringBuilderAsync("xl/styles.xml", sb, cancellationToken).ConfigureAwait(false);
        }

        private bool _completed;
        private Exception? _completionFault;

        /// <summary>
        /// Finalizes the workbook and writes the ZIP footer. No further writes are permitted.
        /// </summary>
        public void Complete()
        {
            if (_completionFault is not null)
                ExceptionDispatchInfo.Capture(_completionFault).Throw();
            if (_completed) return;
            try
            {
                ValidateTables();
                if (_sheetNames.Count == 0)
                    throw new InvalidOperationException("At least one worksheet must be added before completing the workbook.");
                if (_sheetStream is not null)
                    CloseCurrentSheet(_currentSheetIndex);
                WriteWorkbookFinal();
                WriteWorkbookRelsFinal();
                WriteContentTypesFinal();
                WriteStyles();

                if (_useSharedStrings && _sharedStrings.Count > 0)
                    WriteSharedStringsTable();

                WriteTables();

                for (int i = 0; i < _sheetNames.Count; i++)
                    WriteSheetRels(i);

                for (int i = 0; i < _sheetNames.Count; i++)
                {
                    WriteComments(_sheets[i].Comments, i + 1);
                    WriteCommentsVml(_sheets[i].Comments, i + 1);
                }

                WriteDrawingsAndImages();
                _completed = true;
            }
            catch (Exception ex)
            {
                _completionFault = ex;
                throw;
            }
        }

        /// <summary>
        /// Asynchronously finalizes the workbook and writes the ZIP footer. No further writes are permitted.
        /// </summary>
        public async Task CompleteAsync(CancellationToken cancellationToken = default)
        {
            if (_completionFault is not null)
                ExceptionDispatchInfo.Capture(_completionFault).Throw();
            try
            {
                await CompleteAsyncCore(cancellationToken).ConfigureAwait(false);
            }
            catch (Exception ex)
            {
                _completionFault = ex;
                throw;
            }
        }

        private async Task CompleteAsyncCore(CancellationToken cancellationToken)
        {
            if (_completed) return;
            ValidateTables();

            if (_sheetStream is not null)
            {
                await CloseCurrentSheetAsync(_currentSheetIndex, cancellationToken).ConfigureAwait(false);
            }

            if (_sheetNames.Count == 0)
                throw new InvalidOperationException("At least one worksheet must be added before completing the workbook.");
            await WriteWorkbookFinalAsync(cancellationToken).ConfigureAwait(false);
            await WriteWorkbookRelsFinalAsync(cancellationToken).ConfigureAwait(false);
            await WriteContentTypesFinalAsync(cancellationToken).ConfigureAwait(false);
            await WriteStylesAsync(cancellationToken).ConfigureAwait(false);

            if (_useSharedStrings && _sharedStrings.Count > 0)
            {
                await WriteSharedStringsTableAsync(cancellationToken).ConfigureAwait(false);
            }

            await WriteTablesAsync(cancellationToken).ConfigureAwait(false);

            for (int i = 0; i < _sheetNames.Count; i++)
            {
                await WriteSheetRelsAsync(i, cancellationToken).ConfigureAwait(false);
            }

            for (int i = 0; i < _sheetNames.Count; i++)
            {
                await WriteCommentsAsync(_sheets[i].Comments, i + 1, cancellationToken).ConfigureAwait(false);
                await WriteCommentsVmlAsync(_sheets[i].Comments, i + 1, cancellationToken).ConfigureAwait(false);
            }

            await WriteDrawingsAndImagesAsync(cancellationToken).ConfigureAwait(false);

            _completed = true;
        }

        /// <summary>
        /// Finalizes the workbook if needed and releases all resources.
        /// </summary>
        public void Dispose()
        {
            if (_disposed) return;

            try
            {
                if (!_completed && _completionFault is null && _sheetNames.Count > 0) Complete();
            }
            finally
            {
                _disposed = true;
                try { _zip.Dispose(); }
                catch when (_completionFault is not null) { }
                _sink.Dispose();
            }
        }

        /// <summary>
        /// Asynchronously finalizes the workbook if needed and releases all resources.
        /// </summary>
        public async ValueTask DisposeAsync()
        {
            if (_disposed) return;

            try
            {
                if (!_completed && _completionFault is null && _sheetNames.Count > 0) await CompleteAsync().ConfigureAwait(false);
            }
            finally
            {
                _disposed = true;
                try { await _zip.DisposeAsync().ConfigureAwait(false); }
                catch when (_completionFault is not null) { }
                _sink.Dispose();
            }
        }

    }
}
