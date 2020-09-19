using Magicodes.ExporterAndImporter.Core;
using System.Collections.Generic;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    public class Issue131
    {
        public List<DTO_Product> List { get; set; }

    }

    public class DTO_Product
    {
        [ExportImageField(Width = 200, Height = 120)]
        [ExporterHeader("图片")]
        public string ImageUrl { set; get; }

        [ExporterHeader("数量")]
        public int Qty { set; get; }

        [ExporterHeader("价格")]
        public decimal Price { set; get; }

        [ExporterHeader("总额")]
        public decimal TotalPrice => Qty * Price;

        [ExporterHeader()]
        public string Package { set; get; }

        [ExporterHeader()]
        public decimal Size { set; get; }
    }
}
