// ======================================================================
// 
//           Copyright (C) 2019-2030 湖南心莱信息科技有限公司
//           All rights reserved
// 
//           filename : BatchPortraitReceiptInfoInput.cs
//           description :
// 
//           created by 雪雁 at  -- 
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using System.Collections.Generic;
using DinkToPdf;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Pdf;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    /// <summary>
    /// 批量
    /// </summary>
    [PdfExporter(Orientation = Orientation.Portrait, PaperKind = PaperKind.A4)]
    public class BatchPortraitReceiptInfoInput
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
        public List<BatchPortraitReceiptInfoDto> ReceiptInfoInputs { get; set; }
    }
}