using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
    public class XlsxIO_ReadCompare_Benchmarks
    {
        public sealed class ReadRowDto
        {
            public string OrderNo { get; set; } = "";
            public string Region { get; set; } = "";
            public string Product { get; set; } = "";
            public int Qty { get; set; }
        }

        private byte[] _sameFile = Array.Empty<byte>();

        [GlobalSetup]
        public void Setup()
        {
            var data = Enumerable.Range(0, 10_000)
                .Select(i => new ReadRowDto
                {
                    OrderNo = $"O{i}",
                    Region = $"R{i % 8}",
                    Product = $"P{i % 64}",
                    Qty = i % 100,
                })
                .ToArray();

            _sameFile = Xlsx.ToBytes(data);
        }

        [BenchmarkCategory("read-same-file"), Benchmark(Baseline = true)]
        public int Query_Mio_Count()
            => CountRowsMio(_sameFile);

        [BenchmarkCategory("read-same-file"), Benchmark]
        public int Query_MiniExcel_Count()
            => CountRowsMiniExcel(_sameFile);

        [BenchmarkCategory("read-same-file"), Benchmark]
        public int Query_OpenXml_RowScan_Count()
            => CountRowsOpenXml(_sameFile);

        private static int CountRowsMio(byte[] bytes)
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
                if (reader.IsStartElement && reader.ElementType == typeof(Row))
                    count++;
            }

            return count;
        }
    }
}
