using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Reflection;
#if NET461
using TuesPechkin;
using System.Drawing.Printing;
#else
using DinkToPdf;
#endif
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Html;
using System.Text;
using System.Runtime.InteropServices;

namespace Magicodes.ExporterAndImporter.Pdf
{
    /// <summary>
    ///     Pdf导出逻辑
    /// </summary>
    public class PdfExporter : IPdfExporter
    {
        private readonly Lazy<HtmlExporter> _htmlExporter;
        private HtmlExporter HtmlExporter => _htmlExporter.Value;
#if NET461

        private static readonly IConverter PdfConverter = new ThreadSafeConverter(new PdfToolset(
            new WinAnyCPUEmbeddedDeployment(
                new TempFolderDeployment())));

        public PdfExporter()
        {
            _htmlExporter = new Lazy<HtmlExporter>();
        }

#else
        private static readonly SynchronizedConverter PdfConverter = new SynchronizedConverter(new PdfTools());
        public PdfExporter()
        {
            var context = new CustomAssemblyLoadContext();
            // Check the platform and load the appropriate Library
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            {
                var appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                var wkHtmlToPdfPath = Path.Combine(appPath, "runtimes", "linux-x64", "native", "wkhtmltox.so");
                if (!File.Exists(wkHtmlToPdfPath))
                {
                    wkHtmlToPdfPath = Path.Combine(appPath, "wkhtmltox.so");
                }

                context.LoadUnmanagedLibrary(wkHtmlToPdfPath);
            }

            _htmlExporter = new Lazy<HtmlExporter>();
        }

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
            var htmlString = await HtmlExporter.ExportListByTemplate(dataItems, htmlTemplate);
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
            var htmlString = await HtmlExporter.ExportByTemplate(data, htmlTemplate);
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
            var result = await ExportPdf(pdfExporterAttribute, htmlString);
#if NETSTANDARD2_1
            await File.WriteAllBytesAsync(fileName, result);
#else
            File.WriteAllBytes(fileName, result);
#endif

            var fileInfo = new ExportFileInfo(fileName, "application/pdf");
            return await Task.FromResult(fileInfo);
        }

        /// <summary>
        /// 导出到bytes
        /// </summary>
        /// <param name="pdfExporterAttribute"></param>
        /// <param name="htmlString"></param>
        /// <returns></returns>
        private Task<byte[]> ExportPdf(
            PdfExporterAttribute pdfExporterAttribute,
            string htmlString)
        {
            var objSettings = new ObjectSettings
            {
#if !NET461
                HtmlContent = htmlString,
                Encoding = Encoding.UTF8,
                PagesCount = pdfExporterAttribute.IsEnablePagesCount ? true : (bool?)null,
#else
                HtmlText = htmlString,
                CountPages = pdfExporterAttribute.IsEnablePagesCount ? true : (bool?)null,
#endif
                WebSettings = { DefaultEncoding = Encoding.UTF8.BodyName },
            };
            if (pdfExporterAttribute.HeaderSettings != null)
                objSettings.HeaderSettings = pdfExporterAttribute.HeaderSettings;

            if (pdfExporterAttribute.FooterSettings != null)
                objSettings.FooterSettings = pdfExporterAttribute?.FooterSettings;

            var htmlToPdfDocument = new HtmlToPdfDocument
            {
                GlobalSettings =
                {
                    PaperSize = pdfExporterAttribute.PaperKind == PaperKind.Custom
                    ? pdfExporterAttribute.PaperSize : pdfExporterAttribute.PaperKind,
                    Orientation = pdfExporterAttribute.Orientation,
#if !NET461
                    //Out = fileName,
                    ColorMode = ColorMode.Color,
#else
                    ProduceOutline = true,
#endif
                    DocumentTitle = pdfExporterAttribute.Name
                },
                Objects =
                {
                    objSettings
    }
            };

            if (pdfExporterAttribute.MarginSettings != null)
            {
                htmlToPdfDocument.GlobalSettings.Margins = pdfExporterAttribute.MarginSettings;
            }


            var result = PdfConverter.Convert(htmlToPdfDocument);
            return Task.FromResult(result);
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
        /// <summary>
        ///		 获取全局导出定义
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        private static PdfExporterAttribute GetExporterAttribute(Type type)
        {
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
        /// <summary>
        /// 简单实现
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="template"></param>
        /// <returns></returns>
        public async Task<byte[]> ExportBytesByTemplate<T>(T data, string template) where T : class
        {
            var exporterAttribute = GetExporterAttribute<T>();
            var htmlString = await HtmlExporter.ExportByTemplate(data, template);
            return await ExportPdf(exporterAttribute, htmlString);
        }

        /// <summary>
        /// 简单实现
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="template"></param>
        /// <returns></returns>
        public async Task<byte[]> ExportListBytesByTemplate<T>(ICollection<T> data, string template) where T : class
        {
            var exporterAttribute = GetExporterAttribute<T>();
            var htmlString = await HtmlExporter.ExportListByTemplate(data, template);
            return await ExportPdf(exporterAttribute, htmlString);
        }

        /// <summary>
        ///		根据模板导出
        /// </summary>
        /// <param name="data"></param>
        /// <param name="template"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public async Task<byte[]> ExportBytesByTemplate(object data, string template, Type type)
        {
            var exporterAttribute = GetExporterAttribute(type);
            var htmlString = await HtmlExporter.ExportByTemplate(data, template, type);
            return await ExportPdf(exporterAttribute, htmlString);
        }

        /// <summary>
        /// 导出Pdf
        /// </summary>
        /// <param name="data"></param>
        /// <param name="pdfExporterAttribute"></param>
        /// <param name="template"></param>
        /// <returns></returns>
        public async Task<byte[]> ExportListBytesByTemplate<T>(ICollection<T> data, PdfExporterAttribute pdfExporterAttribute, string template) where T : class
        {
            var htmlString = await HtmlExporter.ExportListByTemplate(data, template);
            return await ExportPdf(pdfExporterAttribute, htmlString);
        }

        /// <summary>
        /// 	导出Pdf
        /// </summary>
        /// <param name="data"></param>
        /// <param name="pdfExporterAttribute"></param>
        /// <param name="template"></param>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        /// <exception cref="NotImplementedException"></exception>
        public async Task<byte[]> ExportBytesByTemplate<T>(T data, PdfExporterAttribute pdfExporterAttribute, string template) where T : class
        {
            var htmlString = await HtmlExporter.ExportByTemplate(data, template);
            return await ExportPdf(pdfExporterAttribute, htmlString);
        }
    }
}