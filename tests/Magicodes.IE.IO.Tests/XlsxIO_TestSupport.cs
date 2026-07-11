
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using Magicodes.IE.IO;
using Shouldly;
using Xunit;

namespace Magicodes.IE.IO.Tests
{

    public class OrderDto
    {
        public string OrderNo { get; set; } = "";
        public decimal Amount { get; set; }
        public DateTime CreatedAt { get; set; }
    }

    public class OrderWithFormatDto
    {
        [DisplayFormat(DataFormatString = "0.00")]
        public decimal Amount { get; set; }
        [DisplayFormat(DataFormatString = "yyyy-MM-dd")]
        public DateTime Date { get; set; }
        public string Note { get; set; } = "";
    }

    public class DescribedDto
    {
        [Description("备注")]
        public string Note { get; set; } = "";
    }

    public enum Status
    {
        [Description("待支付")]
        Pending = 0,
        [Description("已支付")]
        Paid = 1,
    }

    public class OrderWithEnumDto
    {
        public string OrderNo { get; set; } = "";
        public Status Status { get; set; }
    }

    public class StringDto { public string Name { get; set; } = ""; }

    public class ULongDto
    {
        public string Name { get; set; } = "";
        public ulong Value { get; set; }
    }

    public enum BigFlag : ulong { None = 0, Huge = 1UL << 63 }

    public class ULongEnumDto
    {
        public string Name { get; set; } = "";
        public BigFlag Flag { get; set; }
    }

    public class ULongEnumAsDoubleDto
    {
        public string Name { get; set; } = "";
        public double Flag { get; set; }
    }

    public class DeclOrderDto
    {
        public string Alpha { get; set; } = "";
        public string Bravo { get; set; } = "";
        public string Charlie { get; set; } = "";
    }

    public class OrderWithAttributesDto
    {
        [ExporterHeader(Name = "订单号", IsIgnore = false)]
        public string OrderNo { get; set; } = "";
        [ExporterHeader(IsIgnore = true)]
        public string Secret { get; set; } = "";
        public decimal Amount { get; set; }
    }

    public record struct TestOrderRecord(string OrderNo, decimal Amount);

    public class OrderWithExporterHeaderDto
    {
        [ExporterHeader(Name = "Order ID", Index = 0)]
        public string OrderNo { get; set; } = "";
        [ExporterHeader(Name = "Total", Index = 1)]
        public decimal Amount { get; set; }
    }

    public class ReadbackOrder
    {
        public string Name { get; set; } = "";
        public int Qty { get; set; }
        public decimal Price { get; set; }
    }

    public class ReceiptDto
    {
        public string Customer { get; set; } = "";
        public List<ReceiptItem> Items { get; set; } = new();
    }
    public record ReceiptItem(string Name, int Qty, decimal Price);

    public class TemplateHolder
    {
        public string CustomerName { get; set; } = "";
        public string OrderNo { get; set; } = "";
        public List<TemplateCell> Items { get; set; } = new();
    }
    public class TemplateCell
    {
        public string Name { get; set; } = "";
        public int Qty { get; set; }
        public decimal Price { get; set; }
    }

    public class Row2 { public string A { get; set; } = ""; public string B { get; set; } = ""; }

    public class BoolDto
    {
        public string Name { get; set; } = "";
        public bool Enabled { get; set; }
    }

    public class OffsetEnumDto
    {
        public DateTimeOffset When { get; set; }
        public Status? OptionalStatus { get; set; }
    }

    public class NullableDto
    {
        public string? Name { get; set; }
        public int? Qty { get; set; }
        public decimal? Price { get; set; }
        public DateTime? CreatedAt { get; set; }
        public bool? Enabled { get; set; }
    }

    [XlsxExportable]
    public class ExportableDto
    {
        public string OrderNo { get; set; } = "";
        public decimal Amount { get; set; }
    }

    [XlsxExportable]
    public class ExportableReorderedDto
    {
        public string A { get; set; } = "";
        public string B { get; set; } = "";
        public string C { get; set; } = "";
    }


    public static class XlsxIO_TestSupport
    {

        public static List<List<string?>> ReadSheet(byte[] bytes)
        {
            using var ms = new MemoryStream(bytes);
            using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
            var entry = zip.GetEntry("xl/worksheets/sheet1.xml");
            entry.ShouldNotBeNull("sheet1.xml 必须存在");
            using var es = entry.Open();
            var doc = new XmlDocument();
            doc.Load(es);
            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            var rows = new List<List<string?>>();
            var rowNodes = doc.SelectNodes("/x:worksheet/x:sheetData/x:row", ns);
            if (rowNodes is null) return rows;
            foreach (XmlNode row in rowNodes)
            {
                var list = new List<string?>();
                var cNodes = row.SelectNodes("x:c", ns);
                if (cNodes is null) { rows.Add(list); continue; }
                foreach (XmlNode c in cNodes)
                {
                    var t = c.SelectSingleNode("x:is/x:t", ns);
                    if (t is not null) { list.Add(t.InnerText); continue; }
                    var v = c.SelectSingleNode("x:v", ns);
                    list.Add(v?.InnerText);
                }
                rows.Add(list);
            }
            return rows;
        }


        public static List<(int StyleId, string? Format)> ReadStyles(byte[] bytes)
        {
            using var ms = new MemoryStream(bytes);
            using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
            var entry = zip.GetEntry("xl/styles.xml");
            if (entry is null) return new();
            using var es = entry.Open();
            var doc = new XmlDocument();
            doc.Load(es);
            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var result = new List<(int, string?)>();
            var xfs = doc.SelectNodes("/x:styleSheet/x:cellXfs/x:xf", ns);
            if (xfs is null) return result;
            foreach (XmlNode xf in xfs)
            {
                var numFmtId = xf.Attributes?["numFmtId"]?.Value;
                if (numFmtId is not null && int.TryParse(numFmtId, out var id) && id > 0)
                    result.Add((id, numFmtId));
            }
            return result;
        }


        public static List<List<string?>> ReadSheet(byte[] bytes, int sheetIndex = 1)
        {
            var path = $"xl/worksheets/sheet{sheetIndex}.xml";
            using var ms = new MemoryStream(bytes);
            using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
            var entry = zip.GetEntry(path);
            entry.ShouldNotBeNull($"{path} 必须存在");
            using var es = entry.Open();
            var doc = new XmlDocument();
            doc.Load(es);
            var ns = new XmlNamespaceManager(doc.NameTable);
            ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var rows = new List<List<string?>>();
            var rowNodes = doc.SelectNodes("/x:worksheet/x:sheetData/x:row", ns);
            if (rowNodes is null) return rows;
            foreach (XmlNode row in rowNodes)
            {
                var list = new List<string?>();
                var cNodes = row.SelectNodes("x:c", ns);
                if (cNodes is null) { rows.Add(list); continue; }
                foreach (XmlNode c in cNodes)
                {
                    var t = c.SelectSingleNode("x:is/x:t", ns);
                    if (t is not null) { list.Add(t.InnerText); continue; }
                    var v = c.SelectSingleNode("x:v", ns);
                    list.Add(v?.InnerText);
                }
                rows.Add(list);
            }
            return rows;
        }



        public static void AssertWellFormedXlsx(byte[] bytes)
        {
            bytes.ShouldNotBeNull();
            bytes.Length.ShouldBeGreaterThan(0);
            using var ms = new MemoryStream(bytes);
            using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
            foreach (var part in new[] { "[Content_Types].xml", "xl/workbook.xml", "xl/worksheets/sheet1.xml", "xl/styles.xml", "xl/_rels/workbook.xml.rels" })
            {
                zip.GetEntry(part).ShouldNotBeNull($"{part} 必须存在");
            }
            foreach (var entry in zip.Entries)
            {
                if (!entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                    && !entry.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) continue;
                using var es = entry.Open();
                using var sr = new StreamReader(es);
                var xml = sr.ReadToEnd();
                AssertWellFormedXml(xml, entry.FullName);
            }
            AssertPackageGraph(zip);
        }


        public static void AssertPackageGraph(ZipArchive zip)
        {
            foreach (var relsEntry in zip.Entries.Where(e => e.FullName.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)))
            {
                string sourcePart = relsEntry.FullName == "_rels/.rels"
                    ? ""
                    : relsEntry.FullName.Substring(0, relsEntry.FullName.Length - ".rels".Length).Replace("/_rels/", "/", StringComparison.Ordinal);
                var doc = new XmlDocument { XmlResolver = null };
                using (var stream = relsEntry.Open()) doc.Load(stream);
                var ns = new XmlNamespaceManager(doc.NameTable);
                ns.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships");
                foreach (XmlElement rel in doc.SelectNodes("/r:Relationships/r:Relationship", ns)!)
                {
                    if (string.Equals(rel.GetAttribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                    var target = rel.GetAttribute("Target").Replace('\\', '/').TrimStart('/');
                    var baseDir = sourcePart.Length == 0 ? "" : sourcePart.Substring(0, sourcePart.LastIndexOf('/') + 1);
                    var combined = NormalizePackagePath(baseDir + target);
                    zip.GetEntry(combined).ShouldNotBeNull($"{relsEntry.FullName} -> {target} ({combined}) 必须存在");
                }
            }
        }

        private static string NormalizePackagePath(string path)
        {
            var parts = path.Split('/');
            var result = new List<string>();
            foreach (var part in parts)
            {
                if (part.Length == 0 || part == ".") continue;
                if (part == "..") { if (result.Count > 0) result.RemoveAt(result.Count - 1); }
                else result.Add(part);
            }
            return string.Join("/", result);
        }

        public static string ReadEntry(byte[] bytes, string path)
        {
            using var ms = new MemoryStream(bytes);
            using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
            var entry = zip.GetEntry(path);
            entry.ShouldNotBeNull();
            using var es = entry.Open();
            using var sr = new StreamReader(es);
            return sr.ReadToEnd();
        }

        public static Dictionary<string, string> ReadCriticalXmlParts(byte[] bytes)
        {
            using var ms = new MemoryStream(bytes);
            using var zip = new ZipArchive(ms, ZipArchiveMode.Read);
            var critical = new[]
            {
                "[Content_Types].xml",
                "xl/workbook.xml",
                "xl/worksheets/sheet1.xml",
                "xl/styles.xml",
                "xl/_rels/workbook.xml.rels"
            };
            var result = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var path in critical)
            {
                var entry = zip.GetEntry(path);
                entry.ShouldNotBeNull($"{path} 必须存在");
                using var es = entry!.Open();
                using var sr = new StreamReader(es);
                result[path] = sr.ReadToEnd();
            }
            return result;
        }

        public static void AssertWellFormedXml(string xml, string partName)
        {
            var doc = new XmlDocument { XmlResolver = null };
            Should.NotThrow(() => doc.LoadXml(xml), $"{partName} 不是 well-formed XML");
        }

        public sealed class Order
        {
            public int A { get; set; }
            public string? B { get; set; }
        }

        public static ColumnMeta[] MakeCols() => new[] { new ColumnMeta("A", "A", null, null, false, 0, 0) };

        public static TypedRowPlan<Order> MakeTypedPlan(string[] propertyNames)
        {
            var cols = new ColumnMeta[propertyNames.Length];
            for (int i = 0; i < propertyNames.Length; i++) cols[i] = new ColumnMeta(propertyNames[i], propertyNames[i], null, null, false, 0, i);
            var props = typeof(Order).GetProperties();
            var getters = new Func<Order, CellValue>[propertyNames.Length];
            for (int i = 0; i < propertyNames.Length; i++)
            {
                var p = Array.Find(props, x => x.Name == propertyNames[i])!;
                getters[i] = o => p.GetValue(o) switch
                {
                    int v => CellValue.FromInteger(v),
                    string s => CellValue.FromString(s),
                    _ => CellValue.Null
                };
            }
            return new TypedRowPlan<Order>(cols, new Func<object?, CellValue>[0], getters, new int[propertyNames.Length], new Action<XlsxWriter.XlsxRowWriter, Order, int>?[propertyNames.Length], hasFormulas: false);
        }

        public sealed class NonSeekableReadStream : Stream
        {
            private readonly Stream _inner;
            public NonSeekableReadStream(Stream inner) => _inner = inner;
            public override bool CanRead => _inner.CanRead;
            public override bool CanSeek => false;
            public override bool CanWrite => false;
            public override long Length => throw new NotSupportedException();
            public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
            public override void Flush() => _inner.Flush();
            public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            protected override void Dispose(bool disposing)
            {
                if (disposing) _inner.Dispose();
                base.Dispose(disposing);
            }
        }

        public sealed class ThrowOnAsyncWriteStream : MemoryStream
        {
            public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
                => Task.FromException(new IOException("injected async write failure"));
        }

        public sealed class TrackingReadStream : Stream
        {
            private readonly Stream _inner;

            public TrackingReadStream(Stream inner) => _inner = inner;

            public int ReadCount { get; private set; }
            public int ReadAsyncCount { get; private set; }
            public int SeekCount { get; private set; }
            public int PositionGetCount { get; private set; }
            public int PositionSetCount { get; private set; }
            public int LengthGetCount { get; private set; }

            public override bool CanRead => _inner.CanRead;
            public override bool CanSeek => _inner.CanSeek;
            public override bool CanWrite => false;

            public override long Length
            {
                get
                {
                    LengthGetCount++;
                    return _inner.Length;
                }
            }

            public override long Position
            {
                get
                {
                    PositionGetCount++;
                    return _inner.Position;
                }
                set
                {
                    PositionSetCount++;
                    _inner.Position = value;
                }
            }

            public override void Flush() => _inner.Flush();
            public override int Read(byte[] buffer, int offset, int count)
            {
                ReadCount++;
                return _inner.Read(buffer, offset, count);
            }
            public override long Seek(long offset, SeekOrigin origin)
            {
                SeekCount++;
                return _inner.Seek(offset, origin);
            }
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
            {
                ReadAsyncCount++;
                return _inner.ReadAsync(buffer, offset, count, cancellationToken);
            }
            protected override void Dispose(bool disposing)
            {
                if (disposing) _inner.Dispose();
                base.Dispose(disposing);
            }
        }
    }
}
