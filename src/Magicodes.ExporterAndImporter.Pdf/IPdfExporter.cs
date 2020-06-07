using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Pdf
{
    /// <summary>
    /// Pdf导出接口
    /// </summary>
    public interface IPdfExporter : IExportListFileByTemplate, IExportFileByTemplate
    {
        /// <summary>
        /// 导出Pdf
        /// </summary>
        /// <param name="pdfExporterAttribute">Pdf导出设置</param>
        /// <param name="htmlString">HTML模板</param>
        /// <returns></returns>
        Task<byte[]> Export(
            PdfExporterAttribute pdfExporterAttribute,
            string htmlString);
    }
}
