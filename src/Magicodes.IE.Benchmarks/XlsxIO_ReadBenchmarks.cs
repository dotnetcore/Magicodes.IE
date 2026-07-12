using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Globalization;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Columns;
using BenchmarkDotNet.Order;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Magicodes.IE.IO;
using MiniExcelLibs;

namespace Magicodes.IE.Benchmarks
{
    [MemoryDiagnoser]
    [Orderer(SummaryOrderPolicy.FastestToSlowest)]
    [CategoriesColumn]
    public class XlsxIO_ReadBenchmarks
    {
        public sealed class ReadRowDto
        {
            public string OrderNo { get; set; } = "";
            public string Region { get; set; } = "";
            public string Product { get; set; } = "";
            public int Qty { get; set; }
        }

        private byte[] _inline10k = Array.Empty<byte>();
        private byte[] _sst10k = Array.Empty<byte>();
        private byte[] _sparse10k = Array.Empty<byte>();

        [GlobalSetup]
        public void Setup()
        {
            var inline = Enumerable.Range(0, 10_000)
                .Select(i => new ReadRowDto
                {
                    OrderNo = $"O{i}",
                    Region = $"R{i % 8}",
                    Product = $"P{i % 64}",
                    Qty = i % 100,
                })
                .ToArray();

            var repeated = Enumerable.Range(0, 10_000)
                .Select(i => new ReadRowDto
                {
                    OrderNo = $"O{i % 32}",
                    Region = $"R{i % 4}",
                    Product = $"P{i % 16}",
                    Qty = i % 100,
                })
                .ToArray();

            _inline10k = Xlsx.ToBytes(inline);
            _sst10k = Xlsx.ToBytes(repeated, p => p.WithAutoSst(true));
            _sparse10k = BuildSparseWorkbook(10_000);
        }

        [BenchmarkCategory("read-10k"), Benchmark(Baseline = true)]
        public int Query_InlineStrings_Count()
            => CountRows(_inline10k);

        [BenchmarkCategory("read-10k"), Benchmark]
        public int Query_SharedStrings_Count()
            => CountRows(_sst10k);

        [BenchmarkCategory("read-10k"), Benchmark]
        public int Query_SparseColumns_Count()
            => CountRows(_sparse10k);

        [BenchmarkCategory("read-10k"), Benchmark]
        public int Query_MiniExcel_Count()
            => CountRowsMiniExcel(_inline10k);

        [BenchmarkCategory("read-10k"), Benchmark]
        public int Query_OpenXml_Count()
            => CountRowsOpenXml(_inline10k);

        private static int CountRows(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes, writable: false);
            var count = 0;
            foreach (var _ in Xlsx.Read<ReadRowDto>(stream))
                count++;
            return count;
        }

        private static int CountRowsMiniExcel(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes, writable: false);
            var count = 0;
            foreach (var _ in MiniExcel.Query<ReadRowDto>(stream))
                count++;
            return count;
        }

        private static int CountRowsOpenXml(byte[] bytes)
        {
            using var stream = new MemoryStream(bytes, writable: false);
            using var doc = SpreadsheetDocument.Open(stream, false);
            var worksheetPart = doc.WorkbookPart!.WorksheetParts.First();
            using var reader = OpenXmlReader.Create(worksheetPart);

            var count = 0;
            while (reader.Read())
            {
                if (reader.ElementType == typeof(Row) && reader.IsStartElement)
                    count++;
            }

            return count;
        }

        private static byte[] BuildSparseWorkbook(int rows)
        {
            using var ms = new MemoryStream();
            using (var doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook, autoSave: true))
            {
                var wbPart = doc.AddWorkbookPart();
                wbPart.Workbook = new Workbook();

                var wsPart = wbPart.AddNewPart<WorksheetPart>();
                var sheetData = new SheetData();
                wsPart.Worksheet = new Worksheet(sheetData);

                var sheets = wbPart.Workbook.AppendChild(new Sheets());
                sheets.Append(new DocumentFormat.OpenXml.Spreadsheet.Sheet
                {
                    Id = wbPart.GetIdOfPart(wsPart),
                    SheetId = 1,
                    Name = "Sparse",
                });

                static Cell InlineCell(string cellRef, string value) => new()
                {
                    CellReference = cellRef,
                    DataType = CellValues.InlineString,
                    InlineString = new InlineString(new Text(value)),
                };

                var header = new Row { RowIndex = 1U };
                header.Append(
                    InlineCell("A1", "OrderNo"),
                    InlineCell("B1", "Region"),
                    InlineCell("C1", "Product"),
                    InlineCell("D1", "Qty"));
                sheetData.Append(header);

                for (int i = 0; i < rows; i++)
                {
                    int rowIndex = i + 2;
                    var row = new Row { RowIndex = (uint)rowIndex };
                    if ((i & 1) == 0)
                    {
                        row.Append(
                            InlineCell($"A{rowIndex}", $"O{i}"),
                            new Cell
                            {
                                CellReference = $"D{rowIndex}",
                                CellValue = new DocumentFormat.OpenXml.Spreadsheet.CellValue((i % 100).ToString(CultureInfo.InvariantCulture)),
                            });
                    }
                    else
                    {
                        row.Append(
                            InlineCell($"B{rowIndex}", $"R{i % 8}"),
                            InlineCell($"C{rowIndex}", $"P{i % 64}"));
                    }
                    sheetData.Append(row);
                }

                wbPart.Workbook.Save();
            }
            return ms.ToArray();
        }
    }
}
