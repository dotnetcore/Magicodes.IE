using System.ComponentModel.DataAnnotations;
using System.Drawing.Printing;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.ExporterAndImporter.Pdf;
using WkHtmlToPdfDotNet;

namespace Magicodes.Benchmarks.Models
{
    /// <summary>
    ///     在Excel导出中，Name将为Sheet名称
    ///     在HTML、Pdf、Word导出中，Name将为标题
    /// </summary>
    [ExcelExporter(Name = "通用导出测试", Author = "雪雁", AutoFitMaxRows = 5000)]
    [ExcelImporter(MaxCount = 50000)]
    [PdfExporter(Orientation = Orientation.Landscape, PaperKind = WkHtmlToPdfDotNet.PaperKind.A4, IsWriteHtml = true, IsEnablePagesCount = false)]
    public class ExportTestData
    {
        /// <summary>
        /// </summary>
        [Display(Name = "列1")]
        [ImporterHeader(Name = "列1")]
        public string Name1 { get; set; }

        [ExporterHeader(DisplayName = "列2")]
        [ImporterHeader(Name = "列2")]
        public string Name2 { get; set; }

        public string Name3 { get; set; }
        public string Name4 { get; set; }
    }
}