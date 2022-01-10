// ======================================================================
// 
//           filename : BatchReceiptInfoInput.cs
//           description :
// 
//           created by 雪雁 at  2019-11-16 13:59
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using Magicodes.ExporterAndImporter.Core;
using System.Collections.Generic;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    /// <summary>
    ///     批量
    /// </summary>
    [Exporter]
    public class BatchReceiptInfoInput
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
        public List<BatchReceiptInfoDto> ReceiptInfoInputs { get; set; }
    }
}