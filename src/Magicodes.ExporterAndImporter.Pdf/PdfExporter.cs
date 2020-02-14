// ======================================================================
// 
//           filename : PdfExporter.cs
//           description :
// 
//           created by 雪雁 at  2019-09-26 14:59
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
#if NET461
using System.Drawing.Imaging;
using TuesPechkin;
#else
using DinkToPdf;
#endif

using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Html;


namespace Magicodes.ExporterAndImporter.Pdf
{
    /// <summary>
    ///     Pdf导出逻辑
    /// </summary>
    public class PdfExporter : IExportListFileByTemplate, IExportFileByTemplate
    {

#if NET461
        private static readonly IConverter PdfConverter = new ThreadSafeConverter(new PdfToolset(
            new WinAnyCPUEmbeddedDeployment(
                new TempFolderDeployment())));
#else
        private static readonly SynchronizedConverter PdfConverter = new SynchronizedConverter(new PdfTools());
        
#endif
        /// <summary>
        ///     根据模板导出列表
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <param name="dataItems"></param>
        /// <param name="htmlTemplate"></param>
        /// <returns></returns>
        public async Task<ExportFileInfo> ExportListByTemplate<T>(string fileName, ICollection<T> dataItems,
            string htmlTemplate = null) where T : class
        {
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("文件名必须填写!", nameof(fileName));

            var exporterAttribute = GetExporterAttribute<T>();
            var exporter = new HtmlExporter();
            var htmlString = await exporter.ExportListByTemplate(dataItems, htmlTemplate);
            if (exporterAttribute.IsWriteHtml)
                File.WriteAllText(fileName + ".html", htmlString);

            return await ExportPdf(fileName, exporterAttribute, htmlString);
        }

        /// <summary>
        ///     根据模板导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <param name="data"></param>
        /// <param name="htmlTemplate"></param>
        /// <returns></returns>
        public async Task<ExportFileInfo> ExportByTemplate<T>(string fileName, T data, string htmlTemplate)
            where T : class
        {
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException("文件名必须填写!", nameof(fileName));

            var exporterAttribute = GetExporterAttribute<T>();
            var exporter = new HtmlExporter();
            var htmlString = await exporter.ExportByTemplate(data, htmlTemplate);
            if (exporterAttribute.IsWriteHtml)
                File.WriteAllText(fileName + ".html", htmlString);

            return await ExportPdf(fileName, exporterAttribute, htmlString);
        }

        /// <summary>
        ///     获取文档转换配置
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="pdfExporterAttribute"></param>
        /// <param name="htmlString"></param>
        /// <returns></returns>
        private async Task<ExportFileInfo> ExportPdf(string fileName,
            PdfExporterAttribute pdfExporterAttribute,
            string htmlString)
        {
            var objSettings = new ObjectSettings
            {
#if !NET461
                        HtmlContent = htmlString,
                        Encoding = pdfExporterAttribute?.Encoding,
                        PagesCount = pdfExporterAttribute.IsEnablePagesCount,
#else
                HtmlText = htmlString,
                CountPages = pdfExporterAttribute.IsEnablePagesCount,
#endif

                WebSettings = { DefaultEncoding = pdfExporterAttribute?.Encoding.BodyName },

                //HeaderSettings = pdfExporterAttribute?.HeaderSettings,
                //FooterSettings = pdfExporterAttribute?.FooterSettings
            };
            if (pdfExporterAttribute?.HeaderSettings != null)
                objSettings.HeaderSettings = pdfExporterAttribute?.HeaderSettings;

            if (pdfExporterAttribute?.FooterSettings != null)
                objSettings.FooterSettings = pdfExporterAttribute?.FooterSettings;

            var htmlToPdfDocument = new HtmlToPdfDocument
            {
                GlobalSettings =
                {
                    PaperSize = pdfExporterAttribute?.PaperKind,
                    Orientation = pdfExporterAttribute?.Orientation,
                    
#if !NET461
                    //Out = fileName,
                    ColorMode = ColorMode.Color,
#else
                    ProduceOutline = true,
#endif
                    DocumentTitle = pdfExporterAttribute?.Name
                },
                Objects =
                {
                    objSettings
                }
            };
            var result = PdfConverter.Convert(htmlToPdfDocument);
#if NETSTANDARD2_1
            await File.WriteAllBytesAsync(fileName, result);
#else
            File.WriteAllBytes(fileName, result);
#endif

            var fileInfo = new ExportFileInfo(fileName, "application/pdf");
            return await Task.FromResult(fileInfo);
        }


        /// <summary>
        ///     获取全局导出定义
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        private static PdfExporterAttribute GetExporterAttribute<T>() where T : class
        {
            var type = typeof(T);
            var exporterTableAttribute = type.GetAttribute<PdfExporterAttribute>(true);
            if (exporterTableAttribute != null)
                return exporterTableAttribute;

            var export = type.GetAttribute<ExporterAttribute>(true) ?? new PdfExporterAttribute();
            return new PdfExporterAttribute
            {
                FontSize = export.FontSize,
                HeaderFontSize = export.HeaderFontSize
            };
        }
    }
}