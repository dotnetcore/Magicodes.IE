// ======================================================================
// 
//           filename : ImportProductDto.cs
//           description :
// 
//           created by 雪雁 at  2019-11-05 20:02
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using Magicodes.ExporterAndImporter.Core;
using System;
using System.ComponentModel.DataAnnotations;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import
{
    /// <summary>
    /// 测试表头位置
    /// </summary>
    [Importer(HeaderRowIndex = 2)]
    public class ImportProductDto
    {
        /// <summary>
        ///     产品名称
        /// </summary>
        [ImporterHeader(Name = "产品名称", Description = "必填")]
        [Required(ErrorMessage = "产品名称是必填的")]
        public string Name { get; set; }

        /// <summary>
        ///     产品代码
        ///     长度验证
        /// </summary>
        [ImporterHeader(Name = "产品代码", Description = "最大长度为20", AutoTrim = false)]
        [MaxLength(20, ErrorMessage = "产品代码最大长度为20（中文算两个字符）")]
        public string Code { get; set; }

        /// <summary>
        ///  测试GUID
        /// </summary>
        public Guid ProductIdTest1 { get; set; }

        public Guid? ProductIdTest2 { get; set; }

        /// <summary>
        ///     产品条码
        /// </summary>
        [ImporterHeader(Name = "产品条码", FixAllSpace = true)]
        [MaxLength(10, ErrorMessage = "产品条码最大长度为10")]
        [RegularExpression(@"^\d*$", ErrorMessage = "产品条码只能是数字")]
        public string BarCode { get; set; }

        /// <summary>
        ///     客户Id
        /// </summary>
        [ImporterHeader(Name = "客户代码", ColumnIndex = 6)]
        public long ClientId { get; set; }

        /// <summary>
        ///     产品型号
        /// </summary>
        [ImporterHeader(Name = "产品型号")]
        public string Model { get; set; }

        /// <summary>
        ///     申报价值
        /// </summary>
        [ImporterHeader(Name = "申报价值")]
        public double DeclareValue { get; set; }

        /// <summary>
        ///     货币单位
        /// </summary>
        [ImporterHeader(Name = "货币单位")]
        public string CurrencyUnit { get; set; }

        /// <summary>
        ///     品牌名称
        /// </summary>
        [ImporterHeader(Name = "品牌名称")]
        public string BrandName { get; set; }

        /// <summary>
        ///     尺寸
        /// </summary>
        [ImporterHeader(Name = "尺寸(长x宽x高)")]
        public string Size { get; set; }

        /// <summary>
        ///     重量（支持不设置ImporterHeader）
        /// </summary>
        //[ImporterHeader(Name = "重量(KG)")]
        [Display(Name = "重量(KG)")]
        public double? Weight { get; set; }

        /// <summary>
        ///     类型
        /// </summary>
        [ImporterHeader(Name = "类型")]
        public ImporterProductType Type { get; set; }

        /// <summary>
        ///     是否行
        /// </summary>
        [ImporterHeader(Name = "是否行")]
        public bool IsOk { get; set; }

        [ImporterHeader(Name = "公式测试",Format ="yyyy-MM-dd")] public DateTime FormulaTest { get; set; }

        /// <summary>
        ///     身份证
        ///     多个错误测试
        /// </summary>
        [ImporterHeader(Name = "身份证")]
        [RegularExpression(@"(^\d{15}$)|(^\d{18}$)|(^\d{17}(\d|X|x)$)", ErrorMessage = "身份证号码无效！")]
        [StringLength(18, ErrorMessage = "身份证长度不得大于18！")]
        public string IdNo { get; set; }

        [Display(Name = "性别")] public string Sex { get; set; }
    }
}