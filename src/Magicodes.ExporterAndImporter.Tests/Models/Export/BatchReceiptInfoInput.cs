using System.Collections.Generic;
using Magicodes.ExporterAndImporter.Core;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    /// <summary>
    /// 批量
    /// </summary>
    [Exporter()]
    public class BatchReceiptInfoInput
    {
        /// <summary>
        /// 标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Logo地址
        /// </summary>
        public string LogoUrl { get; set; }

        /// <summary>
        /// 印章地址
        /// </summary>
        public string SealUrl { get; set; }

        /// <summary>
        /// 收款人
        /// </summary>
        public string Payee { get; set; }

        /// <summary>
        /// 电子收据输入参数
        /// </summary>
        public List<BatchReceiptInfoDto> ReceiptInfoInputs { get; set; }
    }
}
