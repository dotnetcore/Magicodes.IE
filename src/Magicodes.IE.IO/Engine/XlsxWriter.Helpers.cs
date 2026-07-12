
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Magicodes.IE.IO
{
    public sealed partial class XlsxWriter
    {

        private void ValidateTables()
        {
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var displayNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int sheetIndex = 0; sheetIndex < _sheets.Count; sheetIndex++)
            {
                foreach (var table in _sheets[sheetIndex].Tables)
                {
                    ValidateTableDefinition(table);
                    if (!names.Add(table.Name)) throw new ArgumentException($"Duplicate table name '{table.Name}'.");
                    if (!displayNames.Add(table.DisplayName)) throw new ArgumentException($"Duplicate table displayName '{table.DisplayName}'.");
                }
            }
        }

        private static void ValidateTableDefinition(TableDefinition table)
        {
            if (string.IsNullOrWhiteSpace(table.Name)) throw new ArgumentException("Table name cannot be empty.", nameof(table));
            if (string.IsNullOrWhiteSpace(table.DisplayName)) throw new ArgumentException("Table displayName cannot be empty.", nameof(table));
            if (table.Columns.Count == 0) throw new ArgumentException("A table must define at least one column.", nameof(table));
            int colon = table.Ref.IndexOf(':');
            if (colon <= 0 || colon == table.Ref.Length - 1 || table.Ref.IndexOf(':', colon + 1) >= 0)
                throw new ArgumentException($"Invalid table ref '{table.Ref}'.", nameof(table));
            int firstCol = CellRefHelper.CellRefToCol(table.Ref.Substring(0, colon));
            int lastCol = CellRefHelper.CellRefToCol(table.Ref.Substring(colon + 1));
            if (firstCol < 0 || lastCol < 0)
                throw new ArgumentException($"Invalid table ref '{table.Ref}'.", nameof(table));
            int expectedCount = lastCol - firstCol + 1;
            if (expectedCount != table.Columns.Count)
                throw new ArgumentException($"Table ref '{table.Ref}' spans {expectedCount} columns, but {table.Columns.Count} columns were defined.", nameof(table));
            var columnNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var column in table.Columns)
            {
                if (string.IsNullOrWhiteSpace(column.Name)) throw new ArgumentException("Table column name cannot be empty.", nameof(table));
                if (!columnNames.Add(column.Name)) throw new ArgumentException($"Duplicate table column name '{column.Name}'.", nameof(table));
            }
        }

        private void EnsureNotDisposed()
        {
            if (_disposed) throw new ObjectDisposedException(nameof(XlsxWriter));
        }

        private void RecordFault(Exception exception)
        {
            _completionFault ??= exception;
        }

        private static void ValidateCellRange(string value, string parameterName)
        {
            if (!CellRefHelper.IsCellRange(value))
                throw new ArgumentException($"Invalid cell range '{value}'. Expected A1:B2 format.", parameterName);
        }

        private void ValidateSheetName(string name)
        {
            if (name.Length == 0) throw new ArgumentException("Sheet name cannot be empty", nameof(name));
            if (name.Length > MaxSheetNameLength)
                throw new ArgumentException($"Sheet name length {name.Length} exceeds Excel limit of {MaxSheetNameLength} characters", nameof(name));
            if (name[0] == '\'' || name[name.Length - 1] == '\'')
                throw new ArgumentException("Sheet name cannot start or end with single quote", nameof(name));
            if (name.Equals("History", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("Sheet name cannot be the reserved name 'History'", nameof(name));
            foreach (var c in InvalidSheetNameChars)
            {
                if (name.IndexOf(c) >= 0)
                    throw new ArgumentException($"Sheet name contains invalid character '{c}'", nameof(name));
            }
        }

        private void EnsureSheetCapacity(int sheetIdx)
        {
            while (_sheets.Count <= sheetIdx) _sheets.Add(new SheetState());
        }

        private void WriteEntryBytes(string name, byte[] bytes)
        {
            using var es = _zip.OpenEntry(name, _compression);
            es.Write(bytes, 0, bytes.Length);
        }

        private async Task WriteEntryBytesAsync(string name, byte[] bytes, CancellationToken cancellationToken)
        {
            var es = _zip.OpenEntry(name, _compression);
            try
            {
                await es.WriteAsync(bytes, 0, bytes.Length, cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                await ForwardOnlyZipWriter.DisposeEntryStreamAsync(es).ConfigureAwait(false);
            }
        }

        private void WriteEntryString(string name, string content)
        {
            int maxBytes = Encoding.UTF8.GetMaxByteCount(content.Length);
            byte[] buf = System.Buffers.ArrayPool<byte>.Shared.Rent(maxBytes);
            try
            {
                int n = Encoding.UTF8.GetBytes(content, 0, content.Length, buf, 0);
                using var es = _zip.OpenEntry(name, _compression);
                es.Write(buf, 0, n);
            }
            finally
            {
                System.Buffers.ArrayPool<byte>.Shared.Return(buf);
            }
        }

        private async Task WriteEntryStringAsync(string name, string content, CancellationToken cancellationToken)
        {
            int maxBytes = Encoding.UTF8.GetMaxByteCount(content.Length);
            byte[] buf = System.Buffers.ArrayPool<byte>.Shared.Rent(maxBytes);
            try
            {
                int n = Encoding.UTF8.GetBytes(content, 0, content.Length, buf, 0);
                var es = _zip.OpenEntry(name, _compression);
                try
                {
                    await es.WriteAsync(buf, 0, n, cancellationToken).ConfigureAwait(false);
                }
                finally
                {
                    await ForwardOnlyZipWriter.DisposeEntryStreamAsync(es).ConfigureAwait(false);
                }
            }
            finally
            {
                System.Buffers.ArrayPool<byte>.Shared.Return(buf);
            }
        }

        private void WriteEntryStringBuilder(string name, StringBuilder sb)
        {
#if NETSTANDARD2_0
            WriteEntryString(name, sb.ToString());
#else
            using var es = _zip.OpenEntry(name, _compression);
            _sink.WriteUtf8(sb);
            _sink.FlushTo(es);
#endif
        }

        private async Task WriteEntryStringBuilderAsync(string name, StringBuilder sb, CancellationToken cancellationToken)
        {
#if NETSTANDARD2_0
            await WriteEntryStringAsync(name, sb.ToString(), cancellationToken).ConfigureAwait(false);
#else
            await using var es = _zip.OpenEntry(name, _compression);
            _sink.WriteUtf8(sb);
            await _sink.FlushToAsync(es, cancellationToken).ConfigureAwait(false);
#endif
        }

        private async Task WriteEntryFromSinkAsync(string name, CancellationToken cancellationToken)
        {
            var es = _zip.OpenEntry(name, _compression);
            try
            {
                await _sink.FlushToAsync(es, cancellationToken).ConfigureAwait(false);
            }
            finally
            {
                await ForwardOnlyZipWriter.DisposeEntryStreamAsync(es).ConfigureAwait(false);
            }
        }

        private void WriteRels()
        {
            // This relationship is fixed for every workbook; sheet relationships are written later.
            using var es = _zip.OpenEntry("_rels/.rels", _compression);
            _sink.WriteUtf8("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"u8);
            _sink.WriteUtf8("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\r\n"u8);
            _sink.WriteUtf8("  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>\r\n"u8);
            _sink.WriteUtf8("</Relationships>\r\n"u8);
            _sink.FlushTo(es);
        }

        private void BuildWorkbookXml()
        {
            // Workbook metadata is kept small and is emitted only after all sheets are known.
            _sink.WriteUtf8("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"u8);
            _sink.WriteUtf8("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\r\n"u8);
            _sink.WriteUtf8("  <sheets>\r\n"u8);
            for (int i = 0; i < _sheetNames.Count; i++)
            {
                _sink.WriteUtf8("    <sheet name=\""u8);
                _sink.WriteUtf8(XmlHelper.EscapeXmlAttr(_sheetNames[i]));
                _sink.WriteUtf8("\" sheetId=\""u8);
                _sink.WriteInt32(i + 1);
                _sink.WriteUtf8("\" r:id=\"rId"u8);
                _sink.WriteInt32(i + 1);
                _sink.WriteUtf8("\"/>\r\n"u8);
            }
            _sink.WriteUtf8("  </sheets>\r\n"u8);
            if (_namedRanges.Count > 0)
            {
                _sink.WriteUtf8("  <definedNames>\r\n"u8);
                for (int i = 0; i < _namedRanges.Count; i++)
                {
                    var nr = _namedRanges[i];
                    _sink.WriteUtf8("    <definedName name=\""u8);
                    _sink.WriteUtf8(XmlHelper.EscapeXmlAttr(nr.Name));
                    _sink.WriteUtf8("\""u8);
                    if (!string.IsNullOrEmpty(nr.Comment))
                    {
                        _sink.WriteUtf8(" comment=\""u8);
                        _sink.WriteUtf8(XmlHelper.EscapeXmlAttr(nr.Comment!));
                        _sink.WriteUtf8("\""u8);
                    }
                    if (nr.LocalSheetId.HasValue)
                    {
                        _sink.WriteUtf8(" localSheetId=\""u8);
                        _sink.WriteInt32(nr.LocalSheetId.Value);
                        _sink.WriteUtf8("\""u8);
                    }
                    _sink.WriteUtf8(">"u8);
                    _sink.WriteUtf8(XmlHelper.EscapeXmlAttr(nr.Ref));
                    _sink.WriteUtf8("</definedName>\r\n"u8);
                }
                _sink.WriteUtf8("  </definedNames>\r\n"u8);
            }
            _sink.WriteUtf8("</workbook>\r\n"u8);
        }

        private void WriteWorkbookFinal()
        {
            using var es = _zip.OpenEntry("xl/workbook.xml", _compression);
            BuildWorkbookXml();
            _sink.FlushTo(es);
        }

        private Task WriteWorkbookFinalAsync(CancellationToken cancellationToken)
        {
            BuildWorkbookXml();
            return WriteEntryFromSinkAsync("xl/workbook.xml", cancellationToken);
        }

        private void BuildWorkbookRelsXml()
        {
            _sink.WriteUtf8("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"u8);
            _sink.WriteUtf8("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\r\n"u8);
            _sink.WriteUtf8("  <Relationship Id=\"rIdStyles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>\r\n"u8);
            for (int i = 0; i < _sheetNames.Count; i++)
            {
                _sink.WriteUtf8("  <Relationship Id=\"rId"u8);
                _sink.WriteInt32(i + 1);
                _sink.WriteUtf8("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet"u8);
                _sink.WriteInt32(i + 1);
                _sink.WriteUtf8(".xml\"/>\r\n"u8);
            }
            if (_useSharedStrings && _sharedStrings.Count > 0)
            {
                _sink.WriteUtf8("  <Relationship Id=\"rIdSST\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>\r\n"u8);
            }
            _sink.WriteUtf8("</Relationships>\r\n"u8);
        }

        private void WriteWorkbookRelsFinal()
        {
            using var es = _zip.OpenEntry("xl/_rels/workbook.xml.rels", _compression);
            BuildWorkbookRelsXml();
            _sink.FlushTo(es);
        }

        private Task WriteWorkbookRelsFinalAsync(CancellationToken cancellationToken)
        {
            BuildWorkbookRelsXml();
            return WriteEntryFromSinkAsync("xl/_rels/workbook.xml.rels", cancellationToken);
        }

        private void BuildContentTypesXml()
        {
            _sink.WriteUtf8("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"u8);
            _sink.WriteUtf8("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\r\n"u8);
            _sink.WriteUtf8("  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\r\n"u8);
            _sink.WriteUtf8("  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\r\n"u8);
            _sink.WriteUtf8("  <Default Extension=\"png\" ContentType=\"image/png\"/>\r\n"u8);
            _sink.WriteUtf8("  <Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>\r\n"u8);
            _sink.WriteUtf8("  <Default Extension=\"jpg\" ContentType=\"image/jpeg\"/>\r\n"u8);
            _sink.WriteUtf8("  <Default Extension=\"gif\" ContentType=\"image/gif\"/>\r\n"u8);
            _sink.WriteUtf8("  <Default Extension=\"bmp\" ContentType=\"image/bmp\"/>\r\n"u8);
            _sink.WriteUtf8("  <Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/>\r\n"u8);
            for (int i = 0; i < _sheetNames.Count; i++)
            {
                _sink.WriteUtf8("  <Override PartName=\"/xl/worksheets/sheet"u8);
                _sink.WriteInt32(i + 1);
                _sink.WriteUtf8(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\r\n"u8);
            }
            _sink.WriteUtf8("  <Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\r\n"u8);
            _sink.WriteUtf8("  <Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\r\n"u8);
            if (_useSharedStrings && _sharedStrings.Count > 0)
            {
                _sink.WriteUtf8("  <Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>\r\n"u8);
            }
            for (int i = 0; i < _sheets.Count; i++)
            {
                if (_sheets[i].Images.Count > 0)
                {
                    _sink.WriteUtf8("  <Override PartName=\"/xl/drawings/drawing"u8);
                    _sink.WriteInt32(i + 1);
                    _sink.WriteUtf8(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>\r\n"u8);
                }
                if (_sheets[i].Tables.Count > 0)
                {
                    for (int t = 0; t < _sheets[i].Tables.Count; t++)
                    {
                        _sink.WriteUtf8("  <Override PartName=\"/xl/tables/table"u8);
                        _sink.WriteInt32(i + 1);
                        _sink.WriteUtf8("_"u8);
                        _sink.WriteInt32(t + 1);
                        _sink.WriteUtf8(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml\"/>\r\n"u8);
                    }
                }
            }
            for (int i = 0; i < _sheets.Count; i++)
            {
                if (_sheets[i].Comments.Count > 0)
                {
                    _sink.WriteUtf8("  <Override PartName=\"/xl/comments"u8);
                    _sink.WriteInt32(i + 1);
                    _sink.WriteUtf8(".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml\"/>\r\n"u8);
                }
            }
            _sink.WriteUtf8("</Types>\r\n"u8);
        }

        private void WriteContentTypesFinal()
        {
            using var es = _zip.OpenEntry("[Content_Types].xml", _compression);
            BuildContentTypesXml();
            _sink.FlushTo(es);
        }

        private Task WriteContentTypesFinalAsync(CancellationToken cancellationToken)
        {
            BuildContentTypesXml();
            return WriteEntryFromSinkAsync("[Content_Types].xml", cancellationToken);
        }

        private StringBuilder BuildTableXml(TableDefinition tbl, int tableId)
        {
            var sb = new StringBuilder(256);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
            sb.Append("<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"")
              .Append(tableId)
              .Append("\" name=\"")
              .Append(XmlHelper.EscapeXmlAttr(tbl.Name)).Append("\" displayName=\"")
              .Append(XmlHelper.EscapeXmlAttr(tbl.DisplayName ?? tbl.Name)).Append("\" ref=\"")
              .Append(XmlHelper.EscapeXmlAttr(tbl.Ref)).Append("\"");
            if (tbl.ShowTotalsRow)
            {
                sb.Append(" totalsRowShown=\"1\"");
            }
            sb.Append(">\r\n");
            if (tbl.ShowAutoFilter)
            {
                sb.Append("  <autoFilter ref=\"").Append(XmlHelper.EscapeXmlAttr(tbl.Ref)).Append("\"/>\r\n");
            }
            if (tbl.Columns is { Count: > 0 })
            {
                sb.Append("  <tableColumns count=\"").Append(tbl.Columns.Count).Append("\">\r\n");
                for (int c = 0; c < tbl.Columns.Count; c++)
                {
                    sb.Append("    <tableColumn id=\"").Append(c + 1).Append("\" name=\"")
                      .Append(XmlHelper.EscapeXmlAttr(tbl.Columns[c].Name)).Append("\"/>\r\n");
                }
                sb.Append("  </tableColumns>\r\n");
            }
            sb.Append("  <tableStyleInfo name=\"").Append(XmlHelper.EscapeXmlAttr(tbl.TableStyleName))
              .Append("\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\"/>\r\n");
            sb.Append("</table>\r\n");
            return sb;
        }

        private void WriteTables()
        {
            int tableId = 1;
            for (int sheetIdx = 0; sheetIdx < _sheets.Count; sheetIdx++)
            {
                var tables = _sheets[sheetIdx].Tables;
                for (int t = 0; t < tables.Count; t++)
                {
                    WriteEntryStringBuilder($"xl/tables/table{sheetIdx + 1}_{t + 1}.xml", BuildTableXml(tables[t], tableId++));
                }
            }
        }

        private async Task WriteTablesAsync(CancellationToken cancellationToken)
        {
            int tableId = 1;
            for (int sheetIdx = 0; sheetIdx < _sheets.Count; sheetIdx++)
            {
                var tables = _sheets[sheetIdx].Tables;
                for (int t = 0; t < tables.Count; t++)
                {
                    await WriteEntryStringBuilderAsync($"xl/tables/table{sheetIdx + 1}_{t + 1}.xml", BuildTableXml(tables[t], tableId++), cancellationToken).ConfigureAwait(false);
                }
            }
        }

        private StringBuilder BuildDrawingXml(List<ImageAnchor> images, int drawingNo)
        {
            var dwSb = new StringBuilder(256);
            dwSb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
            dwSb.Append("<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\r\n");
            for (int i = 0; i < images.Count; i++)
            {
                var img = images[i];
                dwSb.Append("  <xdr:twoCellAnchor>\r\n");
                dwSb.Append("    <xdr:from><xdr:col>").Append(CellRefToCol(img.FromCell)).Append("</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>").Append(CellRefToRow(img.FromCell)).Append("</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>\r\n");
                dwSb.Append("    <xdr:to><xdr:col>").Append(CellRefToCol(img.ToCell)).Append("</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>").Append(CellRefToRow(img.ToCell)).Append("</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>\r\n");
                dwSb.Append("    <xdr:pic>\r\n");
                dwSb.Append("      <xdr:nvPicPr><xdr:cNvPr id=\"").Append(i + 2).Append("\" name=\"img").Append(i + 1).Append("\"/><xdr:cNvPicPr/></xdr:nvPicPr>\r\n");
                dwSb.Append("      <xdr:blipFill><a:blip r:embed=\"rId").Append(i + 1).Append("\"/></xdr:blipFill>\r\n");
                dwSb.Append("      <xdr:spPr><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr>\r\n");
                dwSb.Append("    </xdr:pic>\r\n");
                dwSb.Append("    <xdr:clientData/>\r\n");
                dwSb.Append("  </xdr:twoCellAnchor>\r\n");
            }
            dwSb.Append("</xdr:wsDr>\r\n");
            return dwSb;
        }

        private StringBuilder BuildDrawingRelsXml(List<ImageAnchor> images, int drawingNo)
        {
            var relSb = new StringBuilder(256);
            relSb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
            relSb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\r\n");
            for (int i = 0; i < images.Count; i++)
            {
                relSb.Append("  <Relationship Id=\"rId").Append(i + 1)
                  .Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/")
                  .Append(XmlHelper.EscapeXmlAttr(images[i].ImageName)).Append("\"/>\r\n");
            }
            relSb.Append("</Relationships>\r\n");
            return relSb;
        }

        private void WriteDrawingsAndImages()
        {
            for (int i = 0; i < _imageBytes.Count; i++)
            {
                var (path, bytes) = _imageBytes[i];
                WriteEntryBytes(path, bytes);
            }
            for (int sheetIdx = 0; sheetIdx < _sheets.Count; sheetIdx++)
            {
                var images = _sheets[sheetIdx].Images;
                if (images.Count == 0) continue;
                int drawingNo = sheetIdx + 1;
                WriteEntryStringBuilder($"xl/drawings/drawing{drawingNo}.xml", BuildDrawingXml(images, drawingNo));
                WriteEntryStringBuilder($"xl/drawings/_rels/drawing{drawingNo}.xml.rels", BuildDrawingRelsXml(images, drawingNo));
            }
        }

        private async Task WriteDrawingsAndImagesAsync(CancellationToken cancellationToken)
        {
            for (int i = 0; i < _imageBytes.Count; i++)
            {
                var (path, bytes) = _imageBytes[i];
                await WriteEntryBytesAsync(path, bytes, cancellationToken).ConfigureAwait(false);
            }
            for (int sheetIdx = 0; sheetIdx < _sheets.Count; sheetIdx++)
            {
                var images = _sheets[sheetIdx].Images;
                if (images.Count == 0) continue;
                int drawingNo = sheetIdx + 1;
                await WriteEntryStringBuilderAsync($"xl/drawings/drawing{drawingNo}.xml", BuildDrawingXml(images, drawingNo), cancellationToken).ConfigureAwait(false);
                await WriteEntryStringBuilderAsync($"xl/drawings/_rels/drawing{drawingNo}.xml.rels", BuildDrawingRelsXml(images, drawingNo), cancellationToken).ConfigureAwait(false);
            }
        }

        private StringBuilder BuildSheetRelsXml(int sheetIdx)
        {
            // A sheet may have several optional relationship parts, so build this part in one pass.
            int sheetNo = sheetIdx + 1;
            var st = _sheets[sheetIdx];
            var sb = new StringBuilder(256);
            bool any = false;

            if (st.Hyperlinks.Count > 0)
            {
                if (!any) { sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"); sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\r\n"); any = true; }
                var hlinks = st.Hyperlinks;
                for (int i = 0; i < hlinks.Count; i++)
                {
                    sb.Append("  <Relationship Id=\"rIdH").Append(i + 1)
                      .Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"")
                      .Append(XmlHelper.EscapeXmlAttr(hlinks[i].Uri)).Append("\" TargetMode=\"External\"/>\r\n");
                }
            }
            if (st.Images.Count > 0)
            {
                if (!any) { sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"); sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\r\n"); any = true; }
                sb.Append("  <Relationship Id=\"rIdImage").Append(sheetNo)
                  .Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing")
                  .Append(sheetNo).Append(".xml\"/>\r\n");
            }
            if (st.Comments.Count > 0)
            {
                if (!any) { sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"); sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\r\n"); any = true; }
                sb.Append("  <Relationship Id=\"rIdComment").Append(sheetNo)
                  .Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments\" Target=\"../comments")
                  .Append(sheetNo).Append(".xml\"/>\r\n");
                sb.Append("  <Relationship Id=\"rIdVml").Append(sheetNo)
                  .Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing\" Target=\"../drawings/vmlDrawing")
                  .Append(sheetNo).Append(".vml\"/>\r\n");
            }
            if (st.Tables.Count > 0)
            {
                if (!any) { sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"); sb.Append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\r\n"); any = true; }
                for (int t = 0; t < st.Tables.Count; t++)
                {
                    sb.Append("  <Relationship Id=\"rIdTable").Append(sheetNo).Append("_").Append(t + 1)
                      .Append("\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/table\" Target=\"../tables/table")
                      .Append(sheetNo).Append("_").Append(t + 1).Append(".xml\"/>\r\n");
                }
            }
            if (any)
            {
                sb.Append("</Relationships>\r\n");
            }
            return sb;
        }

        private void WriteSheetRels(int sheetIdx)
        {
            var sb = BuildSheetRelsXml(sheetIdx);
            if (sb.Length > 0) WriteEntryStringBuilder($"xl/worksheets/_rels/sheet{sheetIdx + 1}.xml.rels", sb);
        }

        private async Task WriteSheetRelsAsync(int sheetIdx, CancellationToken cancellationToken)
        {
            var sb = BuildSheetRelsXml(sheetIdx);
            if (sb.Length > 0) await WriteEntryStringBuilderAsync($"xl/worksheets/_rels/sheet{sheetIdx + 1}.xml.rels", sb, cancellationToken).ConfigureAwait(false);
        }

        private void WriteSheetPr(OutlineSettings? o, PageSetup? ps)
        {
            bool hasOutline = o is not null;
            // fit-to-page only takes effect when <sheetPr><pageSetUpPr fitToPage="1"/> is set;
            // <pageSetup fitToWidth/fitToHeight> alone is ignored by Excel.
            bool fitToPage = ps is not null && (ps.FitToWidth.HasValue || ps.FitToHeight.HasValue);
            if (!hasOutline && !fitToPage) return;
            var sb = new StringBuilder(96);
            sb.Append("<sheetPr>");
            // CT_SheetPr child order: tabColor?, outlinePr?, pageSetUpPr?
            if (hasOutline)
            {
                sb.Append("<outlinePr summaryBelow=\"").Append(o!.SummaryBelow ? "1" : "0")
                  .Append("\" summaryRight=\"").Append(o.SummaryRight ? "1" : "0").Append("\"/>");
            }
            if (fitToPage)
                sb.Append("<pageSetUpPr fitToPage=\"1\"/>");
            sb.Append("</sheetPr>\r\n");
            _sink.WriteUtf8(sb);
        }

        private void WriteDataValidationsToSink(List<DataValidation> validations)
        {
            if (validations.Count == 0) return;
            var sb = new StringBuilder(256);
            sb.Append("<dataValidations count=\"").Append(validations.Count).Append("\">\r\n");
            foreach (var dv in validations)
            {
                sb.Append("  <dataValidation type=\"").Append(DvTypeName(dv.Type)).Append("\"");
                if (dv.Operator.HasValue)
                    sb.Append(" operator=\"").Append(DvOpName(dv.Operator.Value)).Append("\"");
                if (dv.AllowBlank) sb.Append(" allowBlank=\"1\"");
                if (dv.ShowInputMessage) sb.Append(" showInputMessage=\"1\"");
                if (dv.ShowErrorMessage) sb.Append(" showErrorMessage=\"1\"");
                if (dv.PromptTitle is string promptTitle)
                    sb.Append(" promptTitle=\"").Append(XmlHelper.EscapeXmlAttr(promptTitle)).Append("\"");
                if (dv.PromptBody is string promptBody)
                    sb.Append(" prompt=\"").Append(XmlHelper.EscapeXmlAttr(promptBody)).Append("\"");
                sb.Append(" sqref=\"").Append(dv.CellRange).Append("\">\r\n");
                if (!string.IsNullOrEmpty(dv.Formula1))
                    sb.Append("    <formula1>").Append(XmlHelper.EscapeXmlAttr(dv.Formula1!)).Append("</formula1>\r\n");
                if (!string.IsNullOrEmpty(dv.Formula2))
                    sb.Append("    <formula2>").Append(XmlHelper.EscapeXmlAttr(dv.Formula2!)).Append("</formula2>\r\n");
                sb.Append("  </dataValidation>\r\n");
            }
            sb.Append("</dataValidations>\r\n");
            _sink.WriteUtf8(sb);
        }

        private void WriteTablePartsToSink(List<TableDefinition> tables, int sheetNo)
        {
            if (tables.Count == 0) return;
            var sb = new StringBuilder();
            sb.Append("<tableParts count=\"").Append(tables.Count).Append("\">\r\n");
            for (int t = 0; t < tables.Count; t++)
                sb.Append("  <tablePart r:id=\"rIdTable").Append(sheetNo).Append("_").Append(t + 1).Append("\"/>\r\n");
            sb.Append("</tableParts>\r\n");
            _sink.WriteUtf8(sb);
        }

        private void WriteSheetProtection(SheetProtection? p)
        {
            if (p is null) return;
            var sb = new StringBuilder(128);
            sb.Append("<sheetProtection");
            if (!p.Sheet) sb.Append(" sheet=\"0\"");
            if (!p.FormatCells) sb.Append(" formatCells=\"0\"");
            if (!p.FormatColumns) sb.Append(" formatColumns=\"0\"");
            if (!p.FormatRows) sb.Append(" formatRows=\"0\"");
            if (!p.InsertColumns) sb.Append(" insertColumns=\"0\"");
            if (!p.InsertRows) sb.Append(" insertRows=\"0\"");
            if (!p.DeleteColumns) sb.Append(" deleteColumns=\"0\"");
            if (!p.DeleteRows) sb.Append(" deleteRows=\"0\"");
            if (!p.Sort) sb.Append(" sort=\"0\"");
            if (!p.AutoFilter) sb.Append(" autoFilter=\"0\"");
            if (!p.SelectLockedCells) sb.Append(" selectLockedCells=\"0\"");
            if (!p.SelectUnlockedCells) sb.Append(" selectUnlockedCells=\"0\"");
            if (!string.IsNullOrEmpty(p.PasswordHash)) sb.Append(" password=\"").Append(XmlHelper.EscapeXmlAttr(p.PasswordHash!)).Append("\"");
            sb.Append("/>\r\n");
            _sink.WriteUtf8(sb);
        }

        private void WritePageSetup(PageSetup? ps)
        {
            if (ps is null) return;
            var sb = new StringBuilder(128);
            sb.Append("<pageSetup");
            if (ps.Orientation is string orientation) sb.Append(" orientation=\"").Append(XmlHelper.EscapeXmlAttr(orientation)).Append("\"");
            if (ps.PaperSize.HasValue) sb.Append(" paperSize=\"").Append(ps.PaperSize.Value).Append("\"");
            if (ps.Scale.HasValue) sb.Append(" scale=\"").Append(ps.Scale.Value).Append("\"");
            if (ps.FitToWidth.HasValue) sb.Append(" fitToWidth=\"").Append(ps.FitToWidth.Value).Append("\"");
            if (ps.FitToHeight.HasValue) sb.Append(" fitToHeight=\"").Append(ps.FitToHeight.Value).Append("\"");
            if (ps.BlackAndWhite == true) sb.Append(" blackAndWhite=\"1\"");
            sb.Append("/>\r\n");
            _sink.WriteUtf8(sb);
        }

        private void WritePageMargins(PageSetup? ps)
        {
            var margins = ps ?? new PageSetup();
            var sb = new StringBuilder(160);
            sb.Append("<pageMargins left=\"").Append(margins.MarginLeft.ToString(System.Globalization.CultureInfo.InvariantCulture))
              .Append("\" right=\"").Append(margins.MarginRight.ToString(System.Globalization.CultureInfo.InvariantCulture))
              .Append("\" top=\"").Append(margins.MarginTop.ToString(System.Globalization.CultureInfo.InvariantCulture))
              .Append("\" bottom=\"").Append(margins.MarginBottom.ToString(System.Globalization.CultureInfo.InvariantCulture))
              .Append("\" header=\"").Append(margins.MarginHeader.ToString(System.Globalization.CultureInfo.InvariantCulture))
              .Append("\" footer=\"").Append(margins.MarginFooter.ToString(System.Globalization.CultureInfo.InvariantCulture))
              .Append("\"/>\r\n");
            _sink.WriteUtf8(sb);
        }

        private void WriteHeaderFooter(PageSetup? ps)
        {
            if (ps is null || (ps.OddHeader is null && ps.OddFooter is null && ps.EvenHeader is null && ps.EvenFooter is null)) return;
            var sb = new StringBuilder(160);
            if (ps.EvenHeader is not null || ps.EvenFooter is not null)
                sb.Append("<headerFooter differentOddEven=\"1\">");
            else
                sb.Append("<headerFooter>");
            if (ps.OddHeader is string oddHeader) sb.Append("<oddHeader>").Append(XmlHelper.EscapeXmlText(oddHeader)).Append("</oddHeader>");
            if (ps.OddFooter is string oddFooter) sb.Append("<oddFooter>").Append(XmlHelper.EscapeXmlText(oddFooter)).Append("</oddFooter>");
            if (ps.EvenHeader is string evenHeader) sb.Append("<evenHeader>").Append(XmlHelper.EscapeXmlText(evenHeader)).Append("</evenHeader>");
            if (ps.EvenFooter is string evenFooter) sb.Append("<evenFooter>").Append(XmlHelper.EscapeXmlText(evenFooter)).Append("</evenFooter>");
            sb.Append("</headerFooter>\r\n");
            _sink.WriteUtf8(sb);
        }

        private StringBuilder BuildCommentsXml(List<Comment> comments, int sheetNo)
        {
            var sb = new StringBuilder(256);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
            sb.Append("<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\r\n");
            sb.Append("  <authors>\r\n");
            var authorIds = new Dictionary<string, int>(StringComparer.Ordinal);
            int nextId = 0;
            foreach (var c in comments)
            {
                if (!authorIds.ContainsKey(c.Author))
                    authorIds[c.Author] = nextId++;
            }
            foreach (var kv in authorIds)
                sb.Append("    <author>").Append(XmlHelper.EscapeXmlAttr(kv.Key)).Append("</author>\r\n");
            sb.Append("  </authors>\r\n");
            sb.Append("  <commentList>\r\n");
            foreach (var c in comments)
            {
                int authorId = authorIds[c.Author];
                string cellRef = CellRefToA1(c.Row, c.Col);
                sb.Append("    <comment ref=\"").Append(cellRef)
                  .Append("\" authorId=\"").Append(authorId)
                  .Append("\"><text><r><rPr><sz val=\"11\"/></rPr><t xml:space=\"preserve\">")
                  .Append(XmlHelper.EscapeXmlText(c.Text)).Append("</t></r></text></comment>\r\n");
            }
            sb.Append("  </commentList>\r\n");
            sb.Append("</comments>\r\n");
            return sb;
        }

        private void WriteComments(List<Comment> comments, int sheetNo)
        {
            if (comments.Count == 0) return;
            WriteEntryStringBuilder($"xl/comments{sheetNo}.xml", BuildCommentsXml(comments, sheetNo));
        }

        private async Task WriteCommentsAsync(List<Comment> comments, int sheetNo, CancellationToken cancellationToken)
        {
            if (comments.Count == 0) return;
            await WriteEntryStringBuilderAsync($"xl/comments{sheetNo}.xml", BuildCommentsXml(comments, sheetNo), cancellationToken).ConfigureAwait(false);
        }

        private StringBuilder BuildCommentsVmlXml(List<Comment> comments, int sheetNo)
        {
            var sb = new StringBuilder(1024);
            sb.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n");
            sb.Append("<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\">\r\n");
            sb.Append("  <o:shapelayout v:ext=\"edit\">\r\n");
            sb.Append("    <o:idmap v:ext=\"edit\" data=\"1\"/>\r\n");
            sb.Append("  </o:shapelayout>\r\n");
            sb.Append("  <v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\" path=\"m,l,21600r21600,l21600,xe\">\r\n");
            sb.Append("    <v:stroke joinstyle=\"miter\" />\r\n");
            sb.Append("    <v:path gradientshapeok=\"t\" o:connecttype=\"rect\" />\r\n");
            sb.Append("  </v:shapetype>\r\n");

            for (int i = 0; i < comments.Count; i++)
            {
                var c = comments[i];
                sb.Append("  <v:shape id=\"vml").Append(i + 1).Append("\" type=\"#_x0000_t202\" style=\"position:absolute;z-index:1; visibility:hidden\" fillcolor=\"#ffffe1\" insetmode=\"auto\">\r\n");
                sb.Append("    <v:fill color2=\"#ffffe1\" />\r\n");
                sb.Append("    <v:shadow on=\"t\" color=\"black\" obscured=\"t\" />\r\n");
                sb.Append("    <v:path o:connecttype=\"none\" />\r\n");
                sb.Append("    <v:textbox style=\"mso-direction-alt:auto\"><div style=\"text-align:left\" /></v:textbox>\r\n");
                sb.Append("    <x:ClientData ObjectType=\"Note\">\r\n");
                sb.Append("      <x:MoveWithCells />\r\n");
                sb.Append("      <x:SizeWithCells />\r\n");
                sb.Append("      <x:Anchor>").Append(c.Col).Append(", 15, ").Append(c.Row).Append(", 2, ")
                  .Append(c.Col + 2).Append(", 31, ").Append(c.Row + 4).Append(", 1</x:Anchor>\r\n");
                sb.Append("      <x:AutoFill>False</x:AutoFill>\r\n");
                sb.Append("      <x:Row>").Append(c.Row).Append("</x:Row>\r\n");
                sb.Append("      <x:Column>").Append(c.Col).Append("</x:Column>\r\n");
                sb.Append("    </x:ClientData>\r\n");
                sb.Append("  </v:shape>\r\n");
            }

            sb.Append("</xml>\r\n");
            return sb;
        }

        private void WriteCommentsVml(List<Comment> comments, int sheetNo)
        {
            if (comments.Count == 0) return;
            WriteEntryStringBuilder($"xl/drawings/vmlDrawing{sheetNo}.vml", BuildCommentsVmlXml(comments, sheetNo));
        }

        private async Task WriteCommentsVmlAsync(List<Comment> comments, int sheetNo, CancellationToken cancellationToken)
        {
            if (comments.Count == 0) return;
            await WriteEntryStringBuilderAsync($"xl/drawings/vmlDrawing{sheetNo}.vml", BuildCommentsVmlXml(comments, sheetNo), cancellationToken).ConfigureAwait(false);
        }

        private void WriteConditionalFormatting(List<ConditionalFormatting> cfs)
        {
            if (cfs.Count == 0) return;
            var sb = new StringBuilder(256);
            int priority = 1;
            foreach (var cf in cfs)
            {
                sb.Append("<conditionalFormatting sqref=\"").Append(cf.CellRange).Append("\">\r\n");
                foreach (var rule in cf.Rules)
                {
                    sb.Append("  <cfRule type=\"cellIs\" priority=\"").Append(priority++)
                      .Append("\" operator=\"")
                      .Append(CfOpName(rule.Operator)).Append("\">\r\n");
                    sb.Append("    <formula>").Append(XmlHelper.EscapeXmlAttr(rule.Formula1)).Append("</formula>\r\n");
                    if (!string.IsNullOrEmpty(rule.Formula2))
                        sb.Append("    <formula>").Append(XmlHelper.EscapeXmlAttr(rule.Formula2!)).Append("</formula>\r\n");
                    sb.Append("  </cfRule>\r\n");
                }
                sb.Append("</conditionalFormatting>\r\n");
            }
            _sink.WriteUtf8(sb);
        }

        private static int CellRefToCol(string cellRef) => CellRefHelper.CellRefToCol(cellRef);

        private static int CellRefToRow(string cellRef) => CellRefHelper.CellRefToRow(cellRef);

        private static string CellRefToA1(int row, int col)
        {
            if (col < 0) return "A1";
            row += 1;
#if NETSTANDARD2_0
            char[] buf = new char[12];
#else
            Span<char> buf = stackalloc char[12];
#endif
            int p = 0;
            int n = col;
#if NETSTANDARD2_0
            char[] letters = new char[4];
#else
            Span<char> letters = stackalloc char[4];
#endif
            int lp = 0;
            do
            {
                letters[lp++] = (char)('A' + n % 26);
                n = n / 26 - 1;
            } while (n >= 0);
            for (int i = lp - 1; i >= 0; i--) buf[p++] = letters[i];
#if NETSTANDARD2_0
            var rowText = row.ToString(CultureInfo.InvariantCulture);
            rowText.AsSpan().CopyTo(buf.AsSpan(p));
            p += rowText.Length;
            return new string(buf, 0, p);
#else
            if (!row.TryFormat(buf.Slice(p), out int written, provider: CultureInfo.InvariantCulture))
                return "A1";
            p += written;
            return new string(buf.Slice(0, p));
#endif
        }

        private static string DvTypeName(DataValidationType t) => t switch
        {
            DataValidationType.Any => "any",
            DataValidationType.Integer => "whole",
            DataValidationType.Decimal => "decimal",
            DataValidationType.List => "list",
            DataValidationType.Date => "date",
            DataValidationType.Time => "time",
            DataValidationType.TextLength => "textLength",
            DataValidationType.Custom => "custom",
            _ => "any",
        };

        private static string DvOpName(DataValidationOperator op) => op switch
        {
            DataValidationOperator.Between => "between",
            DataValidationOperator.NotBetween => "notBetween",
            DataValidationOperator.Equal => "equal",
            DataValidationOperator.NotEqual => "notEqual",
            DataValidationOperator.LessThan => "lessThan",
            DataValidationOperator.LessThanOrEqual => "lessThanOrEqual",
            DataValidationOperator.GreaterThan => "greaterThan",
            DataValidationOperator.GreaterThanOrEqual => "greaterThanOrEqual",
            _ => "equal",
        };

    }
}
