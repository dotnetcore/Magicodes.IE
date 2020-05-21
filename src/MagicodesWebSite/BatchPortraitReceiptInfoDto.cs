// ======================================================================
// 
//           filename : BatchPortraitReceiptInfoDto.cs
//           description :
// 
//           created by 雪雁 at  2019-11-25 17:04
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using System;

namespace MagicodesWebSite
{
    /// <summary>
    ///     批量导出Dto
    /// </summary>
    public class BatchPortraitReceiptInfoDto
    {
        /// <summary>
        ///     交易时间
        /// </summary>
        public DateTime TradeTime { get; set; }

        /// <summary>
        ///     姓名
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        ///     身份证
        /// </summary>
        public string IdNo { get; set; }

        /// <summary>
        ///     金额
        /// </summary>
        public decimal Amount { get; set; }

        /// <summary>
        ///     支付方式
        /// </summary>
        public string PaymentMethod { get; set; }

        /// <summary>
        ///     交易状态
        /// </summary>
        public string TradeStatus { get; set; }

        /// <summary>
        ///     备注
        /// </summary>
        public string Remark { get; set; }

        /// <summary>
        ///     年级
        /// </summary>
        public string Grade { get; set; }

        /// <summary>
        ///     专业
        /// </summary>
        public string Profession { get; set; }

        /// <summary>
        ///     大写金额
        /// </summary>
        public string UppercaseAmount { get; set; }

        /// <summary>
        ///     编号
        /// </summary>
        public string Code { get; set; }
    }
}