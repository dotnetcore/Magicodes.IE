
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace Magicodes.IE.IO
{
    internal static class XlsxTemplateExporter
    {
        private static readonly Regex TemplateSingleVarRegex = new(@"\{\{([^!#/][^}]*?)\}\}", RegexOptions.Compiled);
        private static readonly Regex TemplateSheetNameRegex = new(@"\{\{!Sheet:Name=([^}]+)\}\}", RegexOptions.Compiled);
        private static readonly Regex TemplateListBlockRegex = new(@"\{\{#([A-Za-z_][A-Za-z0-9_]*)\}\}(.*?)\{\{/\1\}\}", RegexOptions.Compiled | RegexOptions.Singleline);
        private static readonly Regex RowNumberAttributeRegex = new(@"(<row\b[^>]*\br="")(\d+)("")", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex A1ReferenceRegex = new(@"(?<![A-Za-z0-9_])(\$?[A-Za-z]{1,3}\$?)(\d+)", RegexOptions.Compiled);
        private static readonly Regex ReferenceAttributeRegex = new(@"\b(?:r|ref|location)\s*=\s*""([^""]*)""", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex FormulaRegex = new(@"(<f\b[^>]*>)(.*?)(</f>)", RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.Singleline);

        /// <summary>
        /// Replaces template placeholders and writes the output to a new .xlsx file.
        /// </summary>
        public static async Task ExportAsync<T>(string templatePath, string outputPath, T data, CancellationToken cancellationToken = default)
        {
            if (templatePath is null) throw new ArgumentNullException(nameof(templatePath));
            if (outputPath is null) throw new ArgumentNullException(nameof(outputPath));
            using var inStream = File.OpenRead(templatePath);
            using var outStream = File.Create(outputPath);
            await ExportAsync(inStream, outStream, data, cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Replaces template placeholders in a source stream and writes to the output stream.
        /// </summary>
        public static async Task ExportAsync<T>(Stream templateStream, Stream outputStream, T data, CancellationToken cancellationToken = default)
        {
            if (templateStream is null) throw new ArgumentNullException(nameof(templateStream));
            if (outputStream is null) throw new ArgumentNullException(nameof(outputStream));
            cancellationToken.ThrowIfCancellationRequested();

            var readableTemplate = await EnsureSeekableAsync(templateStream, cancellationToken).ConfigureAwait(false);
            try
            {
                var overrides = new Dictionary<string, byte[]>(StringComparer.Ordinal);
                using (var inZip = new ZipArchive(readableTemplate, ZipArchiveMode.Read, leaveOpen: true))
                {
                    foreach (var entry in inZip.Entries)
                    {
                        // Only worksheet, shared-string, and workbook XML can contain template values.
                        if (IsTextEntry(entry.FullName) && NeedsProcessing(entry.FullName))
                        {
                            string xml = await ReadEntryTextAsync(entry, cancellationToken).ConfigureAwait(false);
                            string processed = ProcessTemplateEntry(entry.FullName, xml, data);

                            if (processed != xml)
                                overrides[entry.FullName] = Encoding.UTF8.GetBytes(processed);
                        }
                    }
                }

                readableTemplate.Position = 0;
                using (var outZip = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen: true))
                using (var inZip = new ZipArchive(readableTemplate, ZipArchiveMode.Read, leaveOpen: true))
                {
                    foreach (var inEntry in inZip.Entries)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        var outEntry = outZip.CreateEntry(inEntry.FullName);
                        using var src = inEntry.Open();
                        using var dst = outEntry.Open();
                        if (overrides.TryGetValue(inEntry.FullName, out var replace))
                        {
                            await dst.WriteAsync(replace, 0, replace.Length, cancellationToken).ConfigureAwait(false);
                        }
                        else
                        {
                            await src.CopyToAsync(dst, 81920, cancellationToken).ConfigureAwait(false);
                        }
                    }
                }
            }
            finally
            {
                if (!ReferenceEquals(readableTemplate, templateStream))
                    readableTemplate.Dispose();
            }
        }

        private static async Task<Stream> EnsureSeekableAsync(Stream templateStream, CancellationToken cancellationToken)
        {
            if (!templateStream.CanRead) throw new ArgumentException("templateStream must be readable.", nameof(templateStream));
            if (templateStream.CanSeek)
            {
                templateStream.Position = 0;
                return templateStream;
            }

            var buffer = new MemoryStream();
            await templateStream.CopyToAsync(buffer, 81920, cancellationToken).ConfigureAwait(false);
            buffer.Position = 0;
            return buffer;
        }

        private static bool IsTextEntry(string name) =>
            name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) || name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);

        private static bool NeedsProcessing(string name) =>
            name.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase)
            || name.Equals("xl/sharedStrings.xml", StringComparison.OrdinalIgnoreCase)
            || name.Equals("xl/workbook.xml", StringComparison.OrdinalIgnoreCase);

        private static string ProcessTemplateEntry<T>(string name, string xml, T data)
        {
            var processed = ProcessTemplateSheet(xml, data);
            if (name.Equals("xl/workbook.xml", StringComparison.OrdinalIgnoreCase))
                processed = TemplateSheetNameRegex.Replace(processed, m => XmlHelper.EscapeXmlAttr(m.Groups[1].Value));
            return processed;
        }

        private static string ProcessTemplateSheet<T>(string sheetXml, T data)
        {
            var matches = TemplateListBlockRegex.Matches(sheetXml);
            for (int i = matches.Count - 1; i >= 0; i--)
            {
                var match = matches[i];
                string listName = match.Groups[1].Value;
                string innerTemplate = match.Groups[2].Value;
                var prop = typeof(T).GetProperty(listName);
                var rows = GetRowBounds(innerTemplate);
                int rowSpan = rows.HasValue ? Math.Max(1, rows.Value.Max - rows.Value.Min + 1) : 0;
                int itemCount = 0;
                var replacement = new StringBuilder();

                if (prop?.GetValue(data) is IEnumerable list)
                {
                    int insertionRow = GetMaxRowNumber(sheetXml.Substring(0, match.Index)) + 1;
                    foreach (var item in list)
                    {
                        string itemXml = TemplateSingleVarRegex.Replace(innerTemplate, mm => ReplaceTemplateVar(mm, item, innerTemplate));
                        if (rows.HasValue)
                            itemXml = ShiftRowReferences(itemXml, insertionRow + itemCount * rowSpan - rows.Value.Min);
                        replacement.Append(itemXml);
                        itemCount++;
                    }
                }

                int rowDelta = rowSpan == 0 ? 0 : (itemCount - 1) * rowSpan;
                string prefix = sheetXml.Substring(0, match.Index);
                string suffix = sheetXml.Substring(match.Index + match.Length);
                if (rowDelta != 0) suffix = ShiftRowReferences(suffix, rowDelta);
                sheetXml = prefix + replacement + suffix;
            }

            sheetXml = TemplateSingleVarRegex.Replace(sheetXml, m => ReplaceTemplateVar(m, data, sheetXml));

            return sheetXml;
        }

        private static (int Min, int Max)? GetRowBounds(string xml)
        {
            var matches = RowNumberAttributeRegex.Matches(xml);
            if (matches.Count == 0) return null;

            int min = int.MaxValue;
            int max = 0;
            foreach (Match match in matches)
            {
                int row = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
                min = Math.Min(min, row);
                max = Math.Max(max, row);
            }
            return (min, max);
        }

        private static int GetMaxRowNumber(string xml)
        {
            var bounds = GetRowBounds(xml);
            return bounds?.Max ?? 0;
        }

        private static string ShiftRowReferences(string xml, int delta)
        {
            if (delta == 0 || xml.Length == 0) return xml;

            xml = RowNumberAttributeRegex.Replace(xml, match =>
                match.Groups[1].Value
                + (int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture) + delta).ToString(CultureInfo.InvariantCulture)
                + match.Groups[3].Value);
            xml = ReferenceAttributeRegex.Replace(xml, match =>
                match.Value.Substring(0, match.Groups[1].Index - match.Index)
                + ShiftA1References(match.Groups[1].Value, delta)
                + match.Value.Substring(match.Groups[1].Index - match.Index + match.Groups[1].Length));
            return FormulaRegex.Replace(xml, match =>
                match.Groups[1].Value
                + ShiftFormulaReferences(match.Groups[2].Value, delta)
                + match.Groups[3].Value);
        }

        private static string ShiftA1References(string value, int delta) =>
            A1ReferenceRegex.Replace(value, match =>
                match.Groups[1].Value
                + (int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture) + delta).ToString(CultureInfo.InvariantCulture));

        private static string ShiftFormulaReferences(string formula, int delta)
        {
            var result = new StringBuilder(formula.Length);
            int start = 0;
            for (int i = 0; i < formula.Length; i++)
            {
                if (formula[i] != '"') continue;
                if (i > start)
                    result.Append(ShiftA1References(formula.Substring(start, i - start), delta));

                int end = ++i;
                while (end < formula.Length)
                {
                    if (formula[end] != '"')
                    {
                        end++;
                        continue;
                    }
                    if (end + 1 < formula.Length && formula[end + 1] == '"')
                    {
                        end += 2;
                        continue;
                    }
                    break;
                }

                if (end >= formula.Length)
                {
                    result.Append(formula.Substring(i - 1));
                    return result.ToString();
                }

                result.Append(formula.Substring(i - 1, end - i + 2));
                i = end;
                start = end + 1;
            }

            if (start < formula.Length)
                result.Append(ShiftA1References(formula.Substring(start), delta));
            return result.ToString();
        }

        private static string ReplaceTemplateVar(Match m, object? item, string source)
        {
            if (item is null) return "";
            string name = m.Groups[1].Value.Trim();
            var prop = item.GetType().GetProperty(name, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
            if (prop is null) return m.Value;
            var value = prop.GetValue(item);
            if (value is null) return "";
            string text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? "";
            int lastOpen = source.LastIndexOf('<', m.Index);
            int lastClose = source.LastIndexOf('>', m.Index);
            return lastOpen > lastClose
                ? XmlHelper.EscapeXmlAttr(text)
                : XmlHelper.EscapeXmlText(text);
        }

        private static async Task<string> ReadEntryTextAsync(ZipArchiveEntry entry, CancellationToken cancellationToken)
        {
            using var es = entry.Open();
            using var sr = new StreamReader(es, Encoding.UTF8);
            cancellationToken.ThrowIfCancellationRequested();
            return await sr.ReadToEndAsync().ConfigureAwait(false);
        }
    }
}