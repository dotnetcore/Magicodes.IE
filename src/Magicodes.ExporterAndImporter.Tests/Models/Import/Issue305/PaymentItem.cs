using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import.Issue305
{

    public class PaymentItem
    {
        /// <summary>
        /// 付款时间
        /// </summary>
        public DateTime PaymentDate { get; set; }

        /// <summary>
        /// 付款金额
        /// </summary>
        public decimal PaymentAmount { get; set; }
    }
}
