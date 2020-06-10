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
        /// <param name="data"></param>
        /// <param name="pdfExporterAttribute">Pdf导出设置</param>
        /// <param name="template">模板</param>
        /// <returns></returns>
        Task<byte[]> ExportListBytesByTemplate<T>(ICollection<T> data, PdfExporterAttribute pdfExporterAttribute,
            string template) where T : class;

        /// <summary>
        /// 导出Pdf
        /// </summary>
        /// <param name="data"></param>
        /// <param name="pdfExporterAttribute"></param>
        /// <param name="template"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        Task<byte[]> ExportBytesByTemplate<T>(T data, PdfExporterAttribute pdfExporterAttribute, string template)
            where T : class;
    }
}