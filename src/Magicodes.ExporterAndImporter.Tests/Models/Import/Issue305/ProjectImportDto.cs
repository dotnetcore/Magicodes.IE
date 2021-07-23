using System;
using System.ComponentModel.DataAnnotations;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import.Issue305
{
    [ExporterAndImporter.Core.Importer(HeaderRowIndex = 2)]
    public class ProjectImportDto : ProjectInfoImportDto
    {
        /// <summary>
        /// 付款日期
        /// </summary>
        [Display(Name = "付款日期")]
        public DateTime? PaymentDate { get; set; }

        /// <summary>
        /// 付款金额
        /// </summary>
        [Display(Name = "付款金额")]
        public decimal? PaymentAmount { get; set; }
    }
}
