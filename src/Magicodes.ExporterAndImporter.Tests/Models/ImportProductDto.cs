using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;
using Magicodes.ExporterAndImporter.Core;

namespace Magicodes.ExporterAndImporter.Tests.Models
{
    public class ImportProductDto
    {
        /// <summary>
        /// 产品名称
        /// </summary>
        [ImporterHeader(Name = "产品名称", Description = "必填")]
        [Required(ErrorMessage = "产品名称是必填的")]
        public string Name { get; set; }
        /// <summary>
        /// 产品代码
        /// </summary>
        [ImporterHeader(Name = "产品代码", Description = "最大长度为8", AutoTrim = false)]
        [MaxLength(8, ErrorMessage = "产品代码最大长度为8")]
        public string Code { get; set; }
        /// <summary>
        /// 产品条码
        /// </summary>
        [ImporterHeader(Name = "产品条码", FixAllSpace = true)]
        [MaxLength(10, ErrorMessage = "产品条码最大长度为10")]
        [RegularExpression(@"^\d*$", ErrorMessage = "产品条码只能是数字")]
        public string BarCode { get; set; }
        /// <summary>
        /// 客户Id
        /// </summary>
        [ImporterHeader(Name = "客户代码")]
        public long ClientId { get; set; }
        /// <summary>
        /// 产品型号
        /// </summary>
        [ImporterHeader(Name = "产品型号")]
        public string Model { get; set; }
        /// <summary>
        /// 申报价值
        /// </summary>
        [ImporterHeader(Name = "申报价值")]
        public double DeclareValue { get; set; }
        /// <summary>
        /// 货币单位
        /// </summary>
        [ImporterHeader(Name = "货币单位")]
        public string CurrencyUnit { get; set; }
        /// <summary>
        /// 品牌名称
        /// </summary>
        [ImporterHeader(Name = "品牌名称")]
        public string BrandName { get; set; }
        /// <summary>
        /// 尺寸
        /// </summary>
        [ImporterHeader(Name = "尺寸(长x宽x高)")]
        public string Size { get; set; }
        /// <summary>
        /// 重量（支持不设置ImporterHeader）
        /// </summary>
        //[ImporterHeader(Name = "重量(KG)")]
        [Display(Name = "重量(KG)")]
        public double Weight { get; set; }

        /// <summary>
        /// 类型
        /// </summary>
        [ImporterHeader(Name = "类型")]
        public ImporterProductType Type { get; set; }

        /// <summary>
        /// 是否行
        /// </summary>
        [ImporterHeader(Name = "是否行")]
        public bool IsOk { get; set; }

        [ImporterHeader(Name = "公式测试")]
        public DateTime FormulaTest { get; set; }

        [ImporterHeader(Name = "身份证")]
        public string IdNo { get; set; }

        [Display(Name = "性别")]
        public string Sex { get; set; }
    }
}
