using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Magicodes.IE.IO;
using Shouldly;
using Xunit;

namespace Magicodes.IE.IO.Tests
{
    public partial class XlsxIO_Tests
    {
        [Fact]
        public void WriteWorkbook_DuplicateSheetNames_ThrowsBeforeWritingDuplicate()
        {
            using var output = new MemoryStream();
            using var writer = new XlsxWriter(output);

            writer.AddSheet("Orders");

            Should.Throw<ArgumentException>(() => writer.AddSheet("orders"));
        }

        [Fact]
        public void WriteWorkbook_TypedSheet_AppliesProfileRowFilter()
        {
            var profile = new ExportProfile<OrderDto>()
                .Where(x => x.OrderNo == "keep");

            var bytes = Xlsx.WriteWorkbookToBytes(
                new Sheet<OrderDto>("Orders", new[]
                {
                    new OrderDto { OrderNo = "drop" },
                    new OrderDto { OrderNo = "keep" },
                }, profile));

            var rows = XlsxIO_TestSupport.ReadSheet(bytes);
            rows.Count.ShouldBe(2);
            rows[1][0].ShouldBe("keep");
        }
    }
}
