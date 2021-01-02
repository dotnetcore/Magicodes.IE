// ======================================================================
// 
//           filename : PdfExporterAttribute.cs
//           description :
// 
//           created by 雪雁 at  2019-11-25 15:44
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

#if NET461
using TuesPechkin;
using System.Drawing.Printing;
using static TuesPechkin.GlobalSettings;
#else
using DinkToPdf;
#endif
using Magicodes.ExporterAndImporter.Core;

namespace Magicodes.ExporterAndImporter.Pdf
{
    /// <summary>
    ///     PDF导出特性
    /// </summary>
    public class PdfExporterAttribute : ExporterAttribute
    {
#if !NET461
        /// <summary>
        ///     方向
        /// </summary>
        public Orientation Orientation { get; set; } = Orientation.Landscape;
#else
        /// <summary>
        ///     方向
        /// </summary>
        public PaperOrientation Orientation { get; set; } = PaperOrientation.Landscape;
#endif

        /// <summary>
        ///     纸张类型（默认A4，必须）
        /// </summary>
        public PaperKind PaperKind { get; set; } = PaperKind.A4;

        /// <summary>
        ///     是否启用分页数
        /// </summary>
        public bool IsEnablePagesCount { get; set; }

        /// <summary>
        ///     是否输出HTML模板
        /// </summary>
        public bool IsWriteHtml { get; set; }

        /// <summary>
        ///     头部设置
        /// </summary>
        public HeaderSettings HeaderSettings { get; set; }

        /// <summary>
        ///     底部设置
        /// </summary>
        public FooterSettings FooterSettings { get; set; }

        /// <summary>
        ///     边距设置
        /// </summary>
        public MarginSettings MarginSettings { get; set; }

        /// <summary>
        ///     纸张大小（仅在PaperKind=custom下生效）
        /// </summary>
        public PechkinPaperSize PaperSize { get; set; }
    }
}