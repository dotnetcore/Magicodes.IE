using System.Collections.Generic;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import.Issue305
{
    public class ProjectPaymentDto : ProjectInfoImportDto
    {
        /// <summary>
        /// 支付情况
        /// </summary>
        public List<PaymentItem> PaymentItems { get; set; }
    }
}
