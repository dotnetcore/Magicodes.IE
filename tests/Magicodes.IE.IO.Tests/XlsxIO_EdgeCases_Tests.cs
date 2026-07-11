using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using Magicodes.IE.IO;
using Shouldly;
using Xunit;

namespace Magicodes.IE.IO.Tests
{
    public partial class XlsxIO_Tests
    {
        [Fact]
        public async Task ExportByTemplateAsync_SeekableTemplateStream_DoesNotDisposeCallerStream()
        {
            using var templateMs = new MemoryStream();
            using (var zip = new ZipArchive(templateMs, ZipArchiveMode.Create, leaveOpen: true))
            {
                var entry = zip.CreateEntry("xl/worksheets/sheet1.xml");
                using var es = entry.Open();
                using var sw = new StreamWriter(es);
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                  + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                  + "<sheetData>"
                  + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>客户:{{CustomerName}}</t></is></c></row>"
                  + "</sheetData></worksheet>");
            }

            templateMs.Position = 0;
            using var outputMs = new MemoryStream();
            await Xlsx.ExportByTemplateAsync(templateMs, outputMs, new TemplateHolder { CustomerName = "张三" });

            templateMs.CanRead.ShouldBeTrue();
            templateMs.Position = 0;
            using var zip2 = new ZipArchive(templateMs, ZipArchiveMode.Read, leaveOpen: true);
            zip2.Entries.Count.ShouldBeGreaterThan(0);
        }

        [Fact]
        public async Task ExportByTemplateAsync_NonSeekableTemplateStream_UsesFallbackAndReplacesContent()
        {
            using var templateMs = new MemoryStream();
            using (var zip = new ZipArchive(templateMs, ZipArchiveMode.Create, leaveOpen: true))
            {
                var sheet = zip.CreateEntry("xl/worksheets/sheet1.xml");
                using var es = sheet.Open();
                using var sw = new StreamWriter(es);
                sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                  + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                  + "<sheetData>"
                  + "<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>客户:{{CustomerName}}</t></is></c></row>"
                  + "</sheetData></worksheet>");
            }
            templateMs.Position = 0;

            using var nonSeekable = new XlsxIO_TestSupport.NonSeekableReadStream(templateMs);
            using var outputMs = new MemoryStream();
            await Xlsx.ExportByTemplateAsync(nonSeekable, outputMs, new TemplateHolder { CustomerName = "李四" });

            var xml = XlsxIO_TestSupport.ReadEntry(outputMs.ToArray(), "xl/worksheets/sheet1.xml");
            xml.ShouldContain("客户:李四");
            xml.ShouldNotContain("{{CustomerName}}");
        }

        [Fact]
        public async Task ExportByTemplateAsync_OutputContainsCriticalXmlParts_AndAllAreWellFormed()
        {
            using var templateMs = new MemoryStream();
            using (var zip = new ZipArchive(templateMs, ZipArchiveMode.Create, leaveOpen: true))
            {
                foreach (var path in new[]
                {
                    "[Content_Types].xml",
                    "_rels/.rels",
                    "xl/workbook.xml",
                    "xl/_rels/workbook.xml.rels",
                    "xl/worksheets/sheet1.xml",
                    "xl/worksheets/_rels/sheet1.xml.rels",
                    "xl/styles.xml"
                })
                {
                    var entry = zip.CreateEntry(path);
                    using var es = entry.Open();
                    using var sw = new StreamWriter(es);
                    switch (path)
                    {
                        case "[Content_Types].xml":
                            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"></Types>");
                            break;
                        case "_rels/.rels":
                            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>");
                            break;
                        case "xl/workbook.xml":
                            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets><sheet name=\"S1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>");
                            break;
                        case "xl/_rels/workbook.xml.rels":
                            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/><Relationship Id=\"rIdStyles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/></Relationships>");
                            break;
                        case "xl/worksheets/sheet1.xml":
                            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>客户:{{CustomerName}}</t></is></c></row></sheetData></worksheet>");
                            break;
                        case "xl/worksheets/_rels/sheet1.xml.rels":
                            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"></Relationships>");
                            break;
                        case "xl/styles.xml":
                            sw.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"></styleSheet>");
                            break;
                    }
                }
            }
            templateMs.Position = 0;

            using var outputMs = new MemoryStream();
            await Xlsx.ExportByTemplateAsync(templateMs, outputMs, new TemplateHolder { CustomerName = "王五" });
            var parts = XlsxIO_TestSupport.ReadCriticalXmlParts(outputMs.ToArray());
            foreach (var kv in parts)
            {
                XlsxIO_TestSupport.AssertWellFormedXml(kv.Value, kv.Key);
            }
            parts["xl/worksheets/sheet1.xml"].ShouldContain("王五");
            parts["xl/workbook.xml"].ShouldContain("S1");
        }







        [Fact]
        public async Task LowLevelWriter_WriteRowsAsync_CompleteAsync_ProducesEquivalentXlsxToSync()
        {
            var data = Enumerable.Range(0, 100)
                .Select(i => new StringDto { Name = $"Row{i}" })
                .ToArray();

            (ColumnMeta[] cols, TypedRowPlan<StringDto> plan) BuildPlan()
            {
                var c = new[] { new ColumnMeta("Name", "Name", null, null, false, 0, 0) };
                var getters = new Func<StringDto, CellValue>[] { o => CellValue.FromString(o.Name) };
                var plan = new TypedRowPlan<StringDto>(
                    c, new Func<object?, CellValue>[0], getters, new int[1],
                    new Action<XlsxWriter.XlsxRowWriter, StringDto, int>?[1], hasFormulas: false);
                return (c, plan);
            }

            byte[] Sync()
            {
                var (c, plan) = BuildPlan();
                using var ms = new MemoryStream();
                using (var w = new XlsxWriter(ms))
                {
                    w.AddSheet("S1");
                    w.ResolveColumnStyles(c);
                    w.WriteSheetMeta(c, freezeHeader: true);
                    w.WriteHeader(c);
                    w.WriteRows(data, plan);
                    w.Complete();
                }
                return ms.ToArray();
            }

            async Task<byte[]> Async()
            {
                var (c, plan) = BuildPlan();
                using var ms = new MemoryStream();
                using (var w = new XlsxWriter(ms))
                {
                    w.AddSheet("S1");
                    w.ResolveColumnStyles(c);
                    w.WriteSheetMeta(c, freezeHeader: true);
                    w.WriteHeader(c);
                    await w.WriteRowsAsync(ToAsyncEnumerable(data), plan).ConfigureAwait(false);
                    await w.CompleteAsync().ConfigureAwait(false);
                    await w.DisposeAsync().ConfigureAwait(false);
                }
                return ms.ToArray();
            }

            var syncBytes = Sync();
            var asyncBytes = await Async();

            var syncParts = UnzipParts(syncBytes);
            var asyncParts = UnzipParts(asyncBytes);
            asyncParts.Keys.OrderBy(k => k).SequenceEqual(syncParts.Keys.OrderBy(k => k)).ShouldBeTrue();
            foreach (var key in syncParts.Keys)
            {
                asyncParts[key].SequenceEqual(syncParts[key]).ShouldBeTrue($"part '{key}' 内容在异步链路下不一致");
            }

            var list = Xlsx.Read<StringDto>(new MemoryStream(asyncBytes)).ToList();
            list.Count.ShouldBe(100);
            list[0].Name.ShouldBe("Row0");
            list[99].Name.ShouldBe("Row99");
        }

        [Fact]
        public async Task CompleteAsync_FailureIsFaultedAndCannotResume()
        {
            await using var output = new XlsxIO_TestSupport.ThrowOnAsyncWriteStream();
            await using var writer = new XlsxWriter(output, "Sheet1");
            writer.WriteHeader(XlsxIO_TestSupport.MakeCols());
            var first = await Should.ThrowAsync<IOException>(() => writer.CompleteAsync());
            var second = await Should.ThrowAsync<IOException>(() => writer.CompleteAsync());
            second.Message.ShouldBe(first.Message);
        }

        private static async IAsyncEnumerable<T> ToAsyncEnumerable<T>(IEnumerable<T> source, [EnumeratorCancellation] CancellationToken ct = default)
        {
            foreach (var item in source)
            {
                ct.ThrowIfCancellationRequested();
                await Task.Yield();
                yield return item;
            }
        }

        private static Dictionary<string, byte[]> UnzipParts(byte[] zip)
        {
            var dict = new Dictionary<string, byte[]>(StringComparer.Ordinal);
            using var ms = new MemoryStream(zip);
            using var za = new ZipArchive(ms, ZipArchiveMode.Read);
            foreach (var e in za.Entries)
            {
                using var s = e.Open();
                using var outMs = new MemoryStream();
                s.CopyTo(outMs);
                dict[e.FullName] = outMs.ToArray();
            }
            return dict;
        }
    }
}
