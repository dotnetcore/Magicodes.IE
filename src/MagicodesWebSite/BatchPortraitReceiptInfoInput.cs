// ======================================================================
// 
//           filename : BatchPortraitReceiptInfoInput.cs
//           description :
// 
//           created by 雪雁 at  2019-11-25 17:09
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using System.Collections.Generic;
using Magicodes.ExporterAndImporter.Pdf;
using WkHtmlToPdfDotNet;

namespace MagicodesWebSite
{
    /// <summary>
    ///     批量
    /// </summary>
    [PdfExporter(Orientation = Orientation.Portrait, PaperKind = PaperKind.A4)]
    public class BatchPortraitReceiptInfoInput
    {
        /// <summary>
        ///     标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        ///     Logo地址
        /// </summary>
        public string LogoUrl { get; set; }

        /// <summary>
        ///     印章地址
        /// </summary>
        public string SealUrl { get; set; }

        /// <summary>
        ///     收款人
        /// </summary>
        public string Payee { get; set; }

        /// <summary>
        ///     电子收据输入参数
        /// </summary>
        public List<BatchPortraitReceiptInfoDto> ReceiptInfoInputs { get; set; }
    }
}