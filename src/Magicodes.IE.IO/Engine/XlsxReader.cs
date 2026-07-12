
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace Magicodes.IE.IO
{

    /// <summary>
    /// Advanced streaming reader for xlsx workbooks.
    /// For ordinary reads, use <see cref="Xlsx.Read{T}(Stream, XlsxReadOptions{T}?, Action{XlsxReadErrorInfo}?)"/>.
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class XlsxReader : IDisposable
    {
        private readonly ZipArchive _zip;
        private readonly Stream _sheetStream;
        private readonly XmlReader _xml;
        private readonly Stream? _xlsxStream;
        private readonly bool _leaveOpen;
        private readonly List<string?> _currentRow = new(16);
        private readonly bool _date1904;
        private bool _headerRead;
        private string[] _headers = Array.Empty<string>();
        private string[] _sharedStrings = Array.Empty<string>();

        /// <summary>
        /// Opens a .xlsx workbook positioned at the first readable worksheet.
        /// </summary>
        public XlsxReader(Stream xlsxStream, bool leaveOpen = true)
        {
            if (xlsxStream is null) throw new ArgumentNullException(nameof(xlsxStream));
            _leaveOpen = leaveOpen;
            _xlsxStream = xlsxStream;
            var zip = new ZipArchive(xlsxStream, ZipArchiveMode.Read, leaveOpen: true);
            Stream? sheetStream = null;
            XmlReader? xml = null;
            try
            {
                _date1904 = ReadWorkbookDate1904(zip);
                var entry = ResolveWorksheetEntry(zip)
                    ?? throw new InvalidDataException("No readable worksheet found in xlsx.");
                sheetStream = entry.Open();
                var safeSettings = new XmlReaderSettings
                {
                    IgnoreWhitespace = true,
                    Async = true,
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                };
                xml = XmlReader.Create(sheetStream, safeSettings);
                _zip = zip;
                _sheetStream = sheetStream;
                _xml = xml;

                LoadSharedStrings();
            }
            catch
            {
                xml?.Dispose();
                sheetStream?.Dispose();
                zip.Dispose();
                if (!_leaveOpen) _xlsxStream?.Dispose();
                throw;
            }
        }

        internal bool Date1904 => _date1904;

        private static bool ReadWorkbookDate1904(ZipArchive zip)
        {
            var workbookEntry = zip.GetEntry("xl/workbook.xml");
            if (workbookEntry is null) return false;

            using var stream = workbookEntry.Open();
            using var reader = XmlReader.Create(stream, new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                IgnoreComments = true,
                IgnoreWhitespace = true,
                XmlResolver = null,
            });
            while (reader.Read())
            {
                if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "workbookPr") continue;
                var value = reader.GetAttribute("date1904");
                return value == "1" || string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
            }
            return false;
        }

        private ZipArchiveEntry? ResolveWorksheetEntry(ZipArchive zip)
        {
            var fromWorkbook = ResolveFirstWorksheetEntryFromWorkbook(zip);
            if (fromWorkbook is not null) return fromWorkbook;

            return zip.Entries.FirstOrDefault(static e =>
                e.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)
                && e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                && e.FullName.IndexOf("/_rels/", StringComparison.OrdinalIgnoreCase) < 0);
        }

        private static ZipArchiveEntry? ResolveFirstWorksheetEntryFromWorkbook(ZipArchive zip)
        {
            var workbookEntry = zip.GetEntry("xl/workbook.xml");
            var workbookRelsEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
            if (workbookEntry is null || workbookRelsEntry is null) return null;

            var relTargets = LoadWorkbookRelationshipTargets(workbookRelsEntry);
            if (relTargets.Count == 0) return null;

            using var es = workbookEntry.Open();
            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                IgnoreComments = true,
                IgnoreWhitespace = true,
                XmlResolver = null,
            };
            using var xr = XmlReader.Create(es, settings);
            while (xr.Read())
            {
                if (xr.NodeType != XmlNodeType.Element || xr.LocalName != "sheet") continue;
                var rid = xr.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                if (string.IsNullOrEmpty(rid)) continue;
                if (!relTargets.TryGetValue(rid, out var target)) continue;
                var normalized = NormalizeWorkbookTarget(target);
                if (normalized is null) continue;
                if (!normalized.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)) continue;
                var entry = zip.GetEntry(normalized);
                if (entry is not null) return entry;
            }

            return null;
        }

        private static Dictionary<string, string> LoadWorkbookRelationshipTargets(ZipArchiveEntry workbookRelsEntry)
        {
            using var es = workbookRelsEntry.Open();
            var settings = new XmlReaderSettings
            {
                DtdProcessing = DtdProcessing.Prohibit,
                IgnoreComments = true,
                IgnoreWhitespace = true,
                XmlResolver = null,
            };
            using var xr = XmlReader.Create(es, settings);
            var map = new Dictionary<string, string>(StringComparer.Ordinal);
            while (xr.Read())
            {
                if (xr.NodeType != XmlNodeType.Element || xr.LocalName != "Relationship") continue;
                var id = xr.GetAttribute("Id");
                var target = xr.GetAttribute("Target");
                var mode = xr.GetAttribute("TargetMode");
                if (string.IsNullOrEmpty(id) || string.IsNullOrEmpty(target)) continue;
                if (string.Equals(mode, "External", StringComparison.OrdinalIgnoreCase)) continue;
                map[id] = target;
            }
            return map;
        }

        private static string? NormalizeWorkbookTarget(string? target)
        {
            if (string.IsNullOrWhiteSpace(target)) return null;
            target = target!.Replace('\\', '/');
            if (target.StartsWith("/", StringComparison.Ordinal))
                return target.TrimStart('/');
            if (target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
                return target;
            return "xl/" + target.TrimStart('/');
        }

        private void LoadSharedStrings()
        {
            var entry = _zip.GetEntry("xl/sharedStrings.xml");
            if (entry is null) return;

            using var es = entry.Open();
            var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, IgnoreComments = true, IgnoreWhitespace = false };
            using var xr = XmlReader.Create(es, settings);
            if (xr.ReadToFollowing("sst")) xr.Read();
            var list = new List<string>();
            do
            {
                if (xr.NodeType != XmlNodeType.Element || xr.LocalName != "si") continue;
                list.Add(ReadSharedStringItem(xr));
            }
            while (xr.Read());
            _sharedStrings = list.ToArray();
        }

        private static string ReadSharedStringItem(XmlReader xr)
        {
            if (xr.IsEmptyElement) return string.Empty;

            StringBuilder? sb = null;
            int depth = xr.Depth;
            while (xr.Read())
            {
                if (xr.NodeType == XmlNodeType.EndElement && xr.LocalName == "si" && xr.Depth == depth)
                    break;
                if (xr.NodeType != XmlNodeType.Element) continue;
                if (xr.LocalName == "rPh")
                {
                    int rPhDepth = xr.Depth;
                    if (!xr.IsEmptyElement)
                        while (xr.Read() && !(xr.NodeType == XmlNodeType.EndElement && xr.LocalName == "rPh" && xr.Depth == rPhDepth)) { }
                    continue;
                }
                if (xr.LocalName != "t") continue;
                sb ??= new StringBuilder();
                sb.Append(xr.ReadElementContentAsString());
                if (xr.NodeType == XmlNodeType.EndElement && xr.LocalName == "si" && xr.Depth == depth)
                    break;
            }
            return sb?.ToString() ?? string.Empty;
        }

        /// <summary>
        /// Reads the header row. Subsequent calls return the same result.
        /// </summary>
        public string[] ReadHeader()
        {
            if (_headerRead) return _headers;
            while (_xml.Read())
            {
                if (_xml.NodeType != XmlNodeType.Element || _xml.LocalName != "row") continue;
                ReadRowCells(_currentRow);
                _headers = BuildHeaders(_currentRow);
                if (_headers.Length > _currentRow.Capacity)
                    _currentRow.Capacity = _headers.Length;
                _currentRow.Clear();
                _headerRead = true;
                return _headers;
            }
            _headerRead = true;
            return Array.Empty<string>();
        }

        /// <summary>
        /// Asynchronously reads the header row. Subsequent calls return the same result.
        /// </summary>
        public async Task<string[]> ReadHeaderAsync(CancellationToken cancellationToken = default)
        {
            if (_headerRead) return _headers;
            while (await _xml.ReadAsync().ConfigureAwait(false))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (_xml.NodeType != XmlNodeType.Element || _xml.LocalName != "row") continue;
                await ReadRowCellsAsync(_currentRow, cancellationToken).ConfigureAwait(false);
                _headers = BuildHeaders(_currentRow);
                if (_headers.Length > _currentRow.Capacity)
                    _currentRow.Capacity = _headers.Length;
                _currentRow.Clear();
                _headerRead = true;
                return _headers;
            }
            _headerRead = true;
            return Array.Empty<string>();
        }

        /// <summary>
        /// Reads the next row into a standalone list, or <see langword="null"/> when the sheet is exhausted.
        /// </summary>
        public IReadOnlyList<string?>? ReadNextRow()
        {
            if (!_headerRead) ReadHeader();
            var row = ReadNextRowView();
            if (row is null) return null;

            var copy = new List<string?>(row.Count);
            copy.AddRange(row);
            return copy;
        }

        /// <summary>
        /// Reads the next row as a reused view; its contents are overwritten on the next read.
        /// </summary>
        public IReadOnlyList<string?>? ReadNextRowView()
        {
            if (!_headerRead) ReadHeader();
            // Reuse the row buffer on purpose. Callers that need to retain it should use ReadNextRow.
            while (_xml.Read())
            {
                if (_xml.NodeType != XmlNodeType.Element || _xml.LocalName != "row") continue;
                ReadRowCells(_currentRow);
                return _currentRow;
            }
            return null;
        }

        /// <summary>
        /// Asynchronously reads the next row as a reused view; its contents are overwritten on the next read.
        /// </summary>
        public async ValueTask<IReadOnlyList<string?>?> ReadNextRowViewAsync(CancellationToken cancellationToken = default)
        {
            if (!_headerRead) await ReadHeaderAsync(cancellationToken).ConfigureAwait(false);
            while (await _xml.ReadAsync().ConfigureAwait(false))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (_xml.NodeType != XmlNodeType.Element || _xml.LocalName != "row") continue;
                await ReadRowCellsAsync(_currentRow, cancellationToken).ConfigureAwait(false);
                return _currentRow;
            }
            return null;
        }

        private enum CellKind : byte
        {
            Other = 0,
            Boolean = 1,
            SharedString = 2,
            FormulaString = 3,
            Date = 4,
        }

        private void ReadRowCells(List<string?> buf)
        {
            buf.Clear();
            if (_xml.IsEmptyElement) return;
            CellKind currentType = CellKind.Other;
            int colRef = 0;
            int currentCol = 0;
            while (_xml.Read())
            {
                if (_xml.NodeType == XmlNodeType.EndElement && _xml.LocalName == "row") break;
                if (_xml.NodeType == XmlNodeType.Element && _xml.LocalName == "c")
                {
                    var rAttr = _xml.GetAttribute("r");
                    int thisCol = rAttr is not null ? CellRefToCol(rAttr) : -1;
                    if (thisCol < 0) { thisCol = colRef; colRef++; }
                    else colRef = thisCol + 1;
                    while (buf.Count <= thisCol) buf.Add(null);
                    currentCol = thisCol;
                    currentType = _xml.GetAttribute("t") switch
                    {
                        "b" => CellKind.Boolean,
                        "s" => CellKind.SharedString,
                        "str" => CellKind.FormulaString,
                        "d" => CellKind.Date,
                        _ => CellKind.Other,
                    };
                    continue;
                }
                if (_xml.NodeType == XmlNodeType.Element && _xml.LocalName == "is")
                {
                    string text = ReadInlineStringValue();
                    buf[currentCol] = text;
                    continue;
                }
                if (_xml.NodeType == XmlNodeType.Element && _xml.LocalName == "v")
                {
                    string text = _xml.ReadElementContentAsString();
                    ApplyCellValue(buf, currentCol, currentType, text);
                    continue;
                }
            }
        }

        private async ValueTask ReadRowCellsAsync(List<string?> buf, CancellationToken cancellationToken)
        {
            buf.Clear();
            if (_xml.IsEmptyElement) return;
            CellKind currentType = CellKind.Other;
            int colRef = 0;
            int currentCol = 0;
            while (await _xml.ReadAsync().ConfigureAwait(false))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (_xml.NodeType == XmlNodeType.EndElement && _xml.LocalName == "row") break;
                if (_xml.NodeType == XmlNodeType.Element && _xml.LocalName == "c")
                {
                    var rAttr = _xml.GetAttribute("r");
                    int thisCol = rAttr is not null ? CellRefToCol(rAttr) : -1;
                    if (thisCol < 0) { thisCol = colRef; colRef++; }
                    else colRef = thisCol + 1;
                    while (buf.Count <= thisCol) buf.Add(null);
                    currentCol = thisCol;
                    currentType = _xml.GetAttribute("t") switch
                    {
                        "b" => CellKind.Boolean,
                        "s" => CellKind.SharedString,
                        "str" => CellKind.FormulaString,
                        "d" => CellKind.Date,
                        _ => CellKind.Other,
                    };
                    continue;
                }
                if (_xml.NodeType == XmlNodeType.Element && _xml.LocalName == "is")
                {
                    string text = await ReadInlineStringValueAsync(cancellationToken).ConfigureAwait(false);
                    buf[currentCol] = text;
                    continue;
                }
                if (_xml.NodeType == XmlNodeType.Element && _xml.LocalName == "v")
                {
                    string text = await _xml.ReadElementContentAsStringAsync().ConfigureAwait(false);
                    ApplyCellValue(buf, currentCol, currentType, text);
                    continue;
                }
            }
        }

        private void ApplyCellValue(List<string?> buf, int currentCol, CellKind currentType, string text)
        {
            switch (currentType)
            {
                case CellKind.Boolean:
                    buf[currentCol] = text switch
                    {
                        "1" => "true",
                        "0" => "false",
                        _ => text,
                    };
                    break;
                case CellKind.SharedString:
                    if (int.TryParse(text, out var idx) && idx >= 0 && idx < _sharedStrings.Length)
                        buf[currentCol] = _sharedStrings[idx];
                    else
                        buf[currentCol] = null;
                    break;
                case CellKind.FormulaString:
                    buf[currentCol] = text;
                    break;
                case CellKind.Date:
                    buf[currentCol] = text;
                    break;
                default:
                    buf[currentCol] = text;
                    break;
            }
        }

        private async ValueTask<string> ReadInlineStringValueAsync(CancellationToken cancellationToken)
        {
            if (_xml.IsEmptyElement) return string.Empty;

            StringBuilder? sb = null;
            string? single = null;
            int depth = _xml.Depth;
            while (await _xml.ReadAsync().ConfigureAwait(false))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (_xml.NodeType == XmlNodeType.EndElement && _xml.LocalName == "is" && _xml.Depth == depth)
                    break;
                if (_xml.NodeType != XmlNodeType.Element) continue;
                if (_xml.LocalName == "rPh")
                {
                    int rPhDepth = _xml.Depth;
                    if (!_xml.IsEmptyElement)
                        while (await _xml.ReadAsync().ConfigureAwait(false) && !(_xml.NodeType == XmlNodeType.EndElement && _xml.LocalName == "rPh" && _xml.Depth == rPhDepth))
                        {
                            cancellationToken.ThrowIfCancellationRequested();
                        }
                    continue;
                }
                if (_xml.LocalName != "t") continue;
                var text = await _xml.ReadElementContentAsStringAsync().ConfigureAwait(false);
                if (sb is not null)
                    sb.Append(text);
                else if (single is null)
                    single = text;
                else
                    sb = new StringBuilder(single).Append(text);
                if (_xml.NodeType == XmlNodeType.EndElement && _xml.LocalName == "is" && _xml.Depth == depth)
                    break;
            }
            if (sb is not null) return sb.ToString();
            return single ?? string.Empty;
        }

        private static string[] BuildHeaders(List<string?> row)
        {
            int count = row.Count;
            var headers = new string[count];
            for (int i = 0; i < count; i++)
                headers[i] = row[i] ?? string.Empty;
            return headers;
        }

        private string ReadInlineStringValue()
        {
            if (_xml.IsEmptyElement) return string.Empty;

            // Common case is a single <t> child; capture it directly without allocating a
            // StringBuilder. Only build one if a second <t> appears (rich text runs).
            string? single = null;
            StringBuilder? sb = null;
            int depth = _xml.Depth;
            while (_xml.Read())
            {
                if (_xml.NodeType == XmlNodeType.EndElement && _xml.LocalName == "is" && _xml.Depth == depth)
                    break;
                if (_xml.NodeType != XmlNodeType.Element) continue;
                if (_xml.LocalName == "rPh")
                {
                    int rPhDepth = _xml.Depth;
                    if (!_xml.IsEmptyElement)
                        while (_xml.Read() && !(_xml.NodeType == XmlNodeType.EndElement && _xml.LocalName == "rPh" && _xml.Depth == rPhDepth)) { }
                    continue;
                }
                if (_xml.LocalName != "t") continue;
                var text = _xml.ReadElementContentAsString();
                if (sb is not null)
                    sb.Append(text);
                else if (single is null)
                    single = text;
                else
                    sb = new StringBuilder(single).Append(text);
                if (_xml.NodeType == XmlNodeType.EndElement && _xml.LocalName == "is" && _xml.Depth == depth)
                    break;
            }
            if (sb is not null) return sb.ToString();
            return single ?? string.Empty;
        }

        private static int CellRefToCol(string cellRef) => CellRefHelper.CellRefToCol(cellRef);

        public void Dispose()
        {
            try { _xml.Dispose(); } finally { try { _sheetStream.Dispose(); } finally { try { _zip.Dispose(); } finally { if (!_leaveOpen) _xlsxStream?.Dispose(); } } }
        }

    }

    /// <summary>
    /// Configures how headers and columns map to properties of the target type <typeparamref name="T"/>.
    /// </summary>
    /// <typeparam name="T">The row model type being read.</typeparam>
    public sealed class XlsxReadOptions<T>
    {
        private readonly Dictionary<string, string> _headerToProp = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<int, string> _colIndexToProp = new();
        private readonly List<CellConverter> _converters = new();

        /// <summary>
        /// Gets the custom cell converters registered for this read.
        /// </summary>
        public IList<CellConverter> Converters => _converters;

        /// <summary>
        /// Maps a column, by zero-based index, to a property. Returns this instance for chaining.
        /// </summary>
        public XlsxReadOptions<T> MapColumn(int colIndex, string? propertyName)
        {
            if (colIndex < 0) throw new ArgumentOutOfRangeException(nameof(colIndex));
            if (propertyName is null) return this;
            _colIndexToProp[colIndex] = ResolvePropertyName(propertyName, nameof(propertyName));
            return this;
        }

        /// <summary>
        /// Maps a header name to a property. Returns this instance for chaining.
        /// </summary>
        public XlsxReadOptions<T> MapHeader(string header, string? propertyName)
        {
            if (header is null) throw new ArgumentNullException(nameof(header));
            if (header.Length == 0) throw new ArgumentException("Header cannot be empty.", nameof(header));
            if (propertyName is null) return this;
            _headerToProp[header] = ResolvePropertyName(propertyName, nameof(propertyName));
            return this;
        }

        /// <summary>
        /// Adds a custom cell converter. Returns this instance for chaining.
        /// </summary>
        public XlsxReadOptions<T> WithConverter(CellConverter converter)
        {
            if (converter is null) throw new ArgumentNullException(nameof(converter));
            _converters.Add(converter);
            return this;
        }

        internal IReadOnlyList<CellConverter>? GetConverters() => _converters.Count > 0 ? _converters : null;

        internal string? Resolve(int colIndex, string? header)
        {
            if (_colIndexToProp.TryGetValue(colIndex, out var p)) return p;
            if (header is not null && _headerToProp.TryGetValue(header, out p)) return p;
            return null;
        }

        private static string ResolvePropertyName(string propertyName, string parameterName)
        {
            var generated = XlsxGeneratedTypeMetadataRegistry.TryGet<T>();
            if (generated is not null)
            {
                for (int i = 0; i < generated.Count; i++)
                {
                    var property = generated[i];
                    if (string.Equals(property.Name, propertyName, StringComparison.OrdinalIgnoreCase))
                    {
                        if (property.Setter is null)
                            throw new ArgumentException($"Property '{property.Name}' must be writable.", parameterName);
                        return property.Name;
                    }
                }
            }
            else
            {
                var prop = typeof(T).GetProperty(propertyName, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
                if (prop is not null)
                {
                    if (!prop.CanWrite) throw new ArgumentException($"Property '{prop.Name}' must be writable.", parameterName);
                    return prop.Name;
                }
            }

            throw new ArgumentException($"Property '{propertyName}' was not found on {typeof(T).Name}.", parameterName);
        }
    }

    /// <summary>
    /// Context passed to the error callback when a cell cannot be parsed.
    /// </summary>
    public sealed class XlsxReadErrorInfo
    {

        /// <summary>
        /// Gets the zero-based data row index, excluding the header row.
        /// </summary>
        public int RowIndex { get; init; }

        /// <summary>
        /// Gets the zero-based column index.
        /// </summary>
        public int ColIndex { get; init; }
        /// <summary>
        /// Gets the matched header text.
        /// </summary>
        public string? Header { get; init; }
        /// <summary>
        /// Gets the target property name.
        /// </summary>
        public string? PropertyName { get; init; }
        /// <summary>
        /// Gets the raw cell text.
        /// </summary>
        public string? RawCellValue { get; init; }
        /// <summary>
        /// Gets the name of the target type.
        /// </summary>
        public string? TargetTypeName { get; init; }
        /// <summary>
        /// Gets the exception that caused the parse failure.
        /// </summary>
        public Exception? Exception { get; init; }
    }
}
