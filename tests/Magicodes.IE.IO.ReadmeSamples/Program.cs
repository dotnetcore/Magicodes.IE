using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using Magicodes.IE.IO;

namespace ReadmeSamples;

public sealed class Order
{
    [ExporterHeader(Name = "订单号", Width = 30)]
    public string OrderNo { get; set; } = "";
    [DisplayFormat(DataFormatString = "0.00")]
    public decimal Amount { get; set; }
    public int Qty { get; set; }
    public decimal Price { get; set; }
    public decimal Total { get; set; }
    public System.DateTime CreatedAt { get; set; }
}

public sealed class Item { public string Name { get; set; } = ""; }

public static class Program
{
    private static async IAsyncEnumerable<Order> YieldOrders()
    {
        await Task.Yield();
        yield return new Order { OrderNo = "O-1" };
    }

    public static async Task Main()
    {
        var orders = new[] { new Order { OrderNo = "O-1" } };
        var items = new[] { new Item { Name = "I-1" } };
        using var stream = new MemoryStream();
        _ = Xlsx.ToBytes(orders, p => p.Sheet("订单表")
            .Column(x => x.OrderNo, c => c.WithName("订单号").WithWidth(30))
            .Column(x => x.Amount, c => c.WithFormat("0.00"))
            .Ignore(x => x.CreatedAt)
            .WithFreezeHeader(true));
        _ = Xlsx.WriteWorkbookToBytes(new Sheet<Order>("Orders", orders), new Sheet<Item>("Items", items));
        Xlsx.WriteWorkbook(stream, new Sheet<Order>("Orders", orders), new Sheet<Item>("Items", items));
        stream.Position = 0;
        _ = Xlsx.Read<Order>(stream).ToList();
        stream.Position = 0;
        await foreach (var _ in Xlsx.ReadAsync<Order>(stream)) { }
        stream.SetLength(0);
        await Xlsx.WriteAsync(stream, YieldOrders());
        _ = Xlsx.ToBytes(orders, options: new XlsxWriteOptions { Compression = CompressionLevel.NoCompression, StrictCellReferences = false });
        _ = Xlsx.ToBytes(orders, p => p
            .Column(x => x.Qty, c => c.WithName("数量"))
            .Column(x => x.Price, c => c.WithName("单价"))
            .Column(x => x.Total, c => c.WithName("合计").WithFormula("A{row}*B{row}")));
        _ = Xlsx.ToBytes(orders, p => p.WithAutoSst(true));
        await Xlsx.ExportByTemplateAsync("/tmp/template.xlsx", "/tmp/output.xlsx", orders[0]);
    }
}
