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

using System;
using Magicodes.ExporterAndImporter.Core;
using WkHtmlToPdfDotNet;

namespace Magicodes.ExporterAndImporter.Pdf
{
    /// <summary>
    ///     PDF导出特性
    /// </summary>
    public class PdfExporterAttribute : ExporterAttribute
    {
        private string _pageSizeName = "A4";
        private double _customPageWidth;
        private double _customPageHeight;
        private PdfMeasurementUnit _pageSizeUnit = PdfMeasurementUnit.Millimeters;

        public string DocumentTitle { get; set; }

        public PdfOrientation PageOrientation { get; set; } = PdfOrientation.Landscape;

        public string PageSizeName
        {
            get => _pageSizeName;
            set => _pageSizeName = value;
        }

        public double CustomPageWidth
        {
            get => _customPageWidth;
            set => _customPageWidth = value;
        }

        public double CustomPageHeight
        {
            get => _customPageHeight;
            set => _customPageHeight = value;
        }

        public PdfMeasurementUnit PageSizeUnit
        {
            get => _pageSizeUnit;
            set => _pageSizeUnit = value;
        }

        public PdfPageSize PageSize
        {
            get => BuildPageSize();
            set => SyncFromPageSize(value);
        }

        public PdfHeaderOptions Header { get; set; }

        public PdfFooterOptions Footer { get; set; }

        public PdfMarginOptions Margins { get; set; }

        /// <summary>
        ///     方向
        /// </summary>
        [Obsolete("Use PageOrientation instead.")]
        public Orientation Orientation
        {
            get => PdfWkHtmlCompatibilityMapper.ToWkHtmlOrientation(PageOrientation);
            set => PageOrientation = PdfWkHtmlCompatibilityMapper.ToPdfOrientation(value);
        }

        /// <summary>
        ///     纸张类型（默认A4，必须）
        /// </summary>
        [Obsolete("Use PageSizeName or CustomPageWidth/CustomPageHeight/PageSizeUnit instead.")]
        public PaperKind PaperKind
        {
            get => PdfWkHtmlCompatibilityMapper.ToWkHtmlPaperKind(PageSize);
            set => PageSize = PdfWkHtmlCompatibilityMapper.ToPdfPageSize(value, PageSize?.CustomSize);
        }

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
        [Obsolete("Use Header instead.")]
        public HeaderSettings HeaderSettings
        {
            get => PdfWkHtmlCompatibilityMapper.ToWkHtmlHeaderSettings(Header);
            set => Header = PdfWkHtmlCompatibilityMapper.ToPdfHeaderOptions(value);
        }

        /// <summary>
        ///     底部设置
        /// </summary>
        [Obsolete("Use Footer instead.")]
        public FooterSettings FooterSettings
        {
            get => PdfWkHtmlCompatibilityMapper.ToWkHtmlFooterSettings(Footer);
            set => Footer = PdfWkHtmlCompatibilityMapper.ToPdfFooterOptions(value);
        }

        /// <summary>
        ///     边距设置
        /// </summary>
        [Obsolete("Use Margins instead.")]
        public MarginSettings MarginSettings
        {
            get => PdfWkHtmlCompatibilityMapper.ToWkHtmlMarginSettings(Margins);
            set => Margins = PdfWkHtmlCompatibilityMapper.ToPdfMarginOptions(value);
        }

        /// <summary>
        ///     纸张大小（仅在PaperKind=custom下生效）
        /// </summary>
        [Obsolete("Use PageSizeName or CustomPageWidth/CustomPageHeight/PageSizeUnit instead.")]
        public PechkinPaperSize PaperSize
        {
            get => PdfWkHtmlCompatibilityMapper.ToWkHtmlPaperSize(PageSize);
            set => PageSize = PdfWkHtmlCompatibilityMapper.WithCustomPaperSize(PageSize, value);
        }

        internal PdfExportOptions ToPdfExportOptions()
        {
            return PdfWkHtmlCompatibilityMapper.ToPdfExportOptions(this);
        }

        /// <summary>
        /// 从 PdfPageSize 同步到友好的属性字段。
        /// </summary>
        private void SyncFromPageSize(PdfPageSize pageSize)
        {
            if (pageSize == null)
            {
                _pageSizeName = "A4";
                _customPageWidth = 0;
                _customPageHeight = 0;
                _pageSizeUnit = PdfMeasurementUnit.Millimeters;
                return;
            }

            if (pageSize.IsCustom && pageSize.CustomSize != null)
            {
                _pageSizeName = null;
                _customPageWidth = pageSize.CustomSize.Width;
                _customPageHeight = pageSize.CustomSize.Height;
                _pageSizeUnit = pageSize.CustomSize.Unit;
                return;
            }

            _pageSizeName = string.IsNullOrWhiteSpace(pageSize.StandardName) ? "A4" : pageSize.StandardName;
            _customPageWidth = 0;
            _customPageHeight = 0;
            _pageSizeUnit = PdfMeasurementUnit.Millimeters;
        }

        private PdfPageSize BuildPageSize()
        {
            if (_customPageWidth > 0 && _customPageHeight > 0)
            {
                return PdfPageSize.Custom(new PdfCustomPaperSize
                {
                    Width = _customPageWidth,
                    Height = _customPageHeight,
                    Unit = _pageSizeUnit
                });
            }

            return PdfPageSize.Standard(string.IsNullOrWhiteSpace(_pageSizeName) ? "A4" : _pageSizeName);
        }
    }
}
