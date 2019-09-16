using Magicodes.ExporterAndImporter.Core;
using System;

namespace Magicodes.ExporterAndImporter.Tests
{
    //[ExcelImporter(SheetName = "产品信息")]
    internal class ProductDto
    {
        /// <summary>
        /// 产品名称
        /// </summary>
        [ImporterHeader(Name = "产品名称")]
        public string Name { get; set; }

        /// <summary>
        /// 产品代码
        /// </summary>
        [ImporterHeader(Name = "产品代码")]
        public string Code { get; set; }

        /// <summary>
        /// 产品条码
        /// </summary>
        [ImporterHeader(Name = "产品条码")]
        public string BarCode { get; set; }

        [ImporterHeader(Name = "库存量")]
        public int StoreNum { get; set; }

        [ImporterHeader(Name = "单价")]
        public decimal Price { get; set; }

        [ImporterHeader(Name = "入库日期")]
        public DateTime dateTime { get; set; }
    }
}