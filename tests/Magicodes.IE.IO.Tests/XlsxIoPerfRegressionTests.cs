
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Magicodes.IE.IO;
using Shouldly;
using Xunit;

namespace Magicodes.IE.IO.Tests
{
    public partial class XlsxIO_Tests
    {
        [Fact]
        public void ToBytes_ProducesValidZipReopenableByZipArchive()
        {
            var data = new List<OrderDto>
            {
                new() { OrderNo = "R1", Amount = 12m, CreatedAt = new DateTime(2026, 1, 1) },
                new() { OrderNo = "R2", Amount = 34.5m, CreatedAt = new DateTime(2026, 2, 2) },
            };

            byte[] bytes = Xlsx.ToBytes(data);
            bytes.Length.ShouldBeGreaterThan(0);

            using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
            var names = zip.Entries.Select(e => e.FullName).ToArray();
            names.ShouldContain("xl/worksheets/sheet1.xml");
            names.ShouldContain("xl/workbook.xml");
            names.ShouldContain("xl/styles.xml");
            names.ShouldContain("[Content_Types].xml");

            var sheet = zip.GetEntry("xl/worksheets/sheet1.xml")!;
            using var sr = new StreamReader(sheet.Open());
            var sheetXml = sr.ReadToEnd();
            sheetXml.ShouldContain("R1");
            sheetXml.ShouldContain("R2");
        }

        [Fact]
        public void ToBytes_RoundTripsViaRead()
        {
            var data = new List<OrderDto>
            {
                new() { OrderNo = "RT-1", Amount = 9.99m, CreatedAt = new DateTime(2026, 3, 3) },
                new() { OrderNo = "RT-2", Amount = 0m, CreatedAt = new DateTime(2026, 4, 4) },
            };

            byte[] bytes = Xlsx.ToBytes(data);
            using var ms = new MemoryStream(bytes);
            var read = Xlsx.Read<OrderDto>(ms).ToList();
            read.Count.ShouldBe(2);
            read[0].OrderNo.ShouldBe("RT-1");
            read[0].Amount.ShouldBe(9.99m);
            read[1].OrderNo.ShouldBe("RT-2");
            read[1].Amount.ShouldBe(0m);
        }

        [Fact]
        public void WriteWorkbookToBytes_ProducesMultipleSheets()
        {
            var s1 = new Sheet("Sheet1", new[] { new OrderDto { OrderNo = "A", Amount = 1 } });
            var s2 = new Sheet("Sheet2", new[] { new OrderDto { OrderNo = "B", Amount = 2 } });
            byte[] bytes = Xlsx.WriteWorkbookToBytes(s1, s2);

            using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
            var names = zip.Entries.Select(e => e.FullName).ToArray();
            names.ShouldContain("xl/worksheets/sheet1.xml");
            names.ShouldContain("xl/worksheets/sheet2.xml");
        }

        [Fact]
        public void ToBytes_WithAutoSst_WritesSharedStringsAndRoundTrips()
        {
            var data = new List<OrderDto>();
            for (int i = 0; i < 20; i++)
                data.Add(new() { OrderNo = "SST1", Amount = i });

            byte[] bytes = Xlsx.ToBytes(data, cfg => cfg.WithAutoSst());
            using var zip = new ZipArchive(new MemoryStream(bytes), ZipArchiveMode.Read);
            zip.Entries.Select(e => e.FullName).ShouldContain("xl/sharedStrings.xml");

            using var ms = new MemoryStream(bytes);
            var read = Xlsx.Read<OrderDto>(ms).ToList();
            read.Count.ShouldBe(20);
            read[0].OrderNo.ShouldBe("SST1");
            read[19].OrderNo.ShouldBe("SST1");
        }
    }
}
