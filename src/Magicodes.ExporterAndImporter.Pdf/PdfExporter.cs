using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;
using System.Reflection;
using WkHtmlToPdfDotNet;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Html;
using Magicodes.IE.Core;

namespace Magicodes.ExporterAndImporter.Pdf
{
    /// <summary>
    ///     Pdf导出逻辑
    /// </summary>
    public class PdfExporter : IPdfExporter
    {
        private readonly Lazy<HtmlExporter> _htmlExporter;
        private readonly IPdfNativeLibraryService _nativeLibraryService;
        private static readonly Lazy<SynchronizedConverter> PdfConverter = new Lazy<SynchronizedConverter>(() => new SynchronizedConverter(new PdfTools()));
        private HtmlExporter HtmlExporter => _htmlExporter.Value;

        /// <summary>
        /// 触发 PdfNativeLibraryBootstrapper 的静态构造函数，注册 DllImportResolver。
        /// 必须在 PdfConverter（使用 Haukcode P/Invoke）之前完成。
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static void EnsureBootstrap() { _ = PdfNativeLibraryBootstrapper.CheckEnvironment(); }

        public PdfExporter() : this(new PdfNativeLibraryService())
        {
        }

        public PdfExporter(IPdfNativeLibraryService nativeLibraryService)
        {
            _htmlExporter = new Lazy<HtmlExporter>();
            _nativeLibraryService = nativeLibraryService ?? throw new ArgumentNullException(nameof(nativeLibraryService));
        }

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
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException(Resource.FileNameMustBeFilled, nameof(fileName));

            var exporterAttribute = GetExporterAttribute<T>();
            var htmlString = await HtmlExporter.ExportListByTemplate(dataItems, htmlTemplate);
            var pdfExportOptions = exporterAttribute.ToPdfExportOptions();
            if (pdfExportOptions.WriteHtml)
                File.WriteAllText(fileName + ".html", htmlString);

            return await ExportPdf(fileName, pdfExportOptions, htmlString);
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
            if (string.IsNullOrWhiteSpace(fileName)) throw new ArgumentException(Resource.FileNameMustBeFilled, nameof(fileName));

            var exporterAttribute = GetExporterAttribute<T>();
            var htmlString = await HtmlExporter.ExportByTemplate(data, htmlTemplate);
            var pdfExportOptions = exporterAttribute.ToPdfExportOptions();
            if (pdfExportOptions.WriteHtml)
                File.WriteAllText(fileName + ".html", htmlString);

            return await ExportPdf(fileName, pdfExportOptions, htmlString);
        }

        /// <summary>
        ///     获取文档转换配置
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="pdfExportOptions"></param>
        /// <param name="htmlString"></param>
        /// <returns></returns>
        private async Task<ExportFileInfo> ExportPdf(string fileName,
            PdfExportOptions pdfExportOptions,
            string htmlString)
        {
            var result = await ExportPdf(pdfExportOptions, htmlString);
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
        /// <param name="pdfExportOptions"></param>
        /// <param name="htmlString"></param>
        /// <returns></returns>
        /// <exception cref="InvalidOperationException">当 wkhtmltopdf native 库不可用时，包含环境诊断信息</exception>
        private Task<byte[]> ExportPdf(
            PdfExportOptions pdfExportOptions,
            string htmlString)
        {
            try
            {
                EnsureBootstrap();
                var htmlToPdfDocument = PdfWkHtmlCompatibilityMapper.ToHtmlToPdfDocument(pdfExportOptions, htmlString);
                var result = PdfConverter.Value.Convert(htmlToPdfDocument);
                return Task.FromResult(result);
            }
            catch (Exception ex) when (IsNativeLibraryError(ex))
            {
                var env = _nativeLibraryService.CheckEnvironment();
                throw new InvalidOperationException(
                    $"PDF export failed: wkhtmltopdf native library could not be loaded.\n\n{env}", ex);
            }
        }

        private static bool IsNativeLibraryError(Exception ex)
        {
            var current = ex;
            while (current != null)
            {
                if (current is DllNotFoundException ||
                    current is BadImageFormatException ||
                    current is EntryPointNotFoundException ||
                    current is PlatformNotSupportedException)
                    return true;
                if (current is NotSupportedException &&
                    current.Message != null &&
                    current.Message.IndexOf("native library", StringComparison.OrdinalIgnoreCase) >= 0)
                    return true;
                current = current.InnerException;
            }
            return false;
        }

        /// <summary>
        ///     获取全局导出定义
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        private static PdfExporterAttribute GetExporterAttribute<T>() where T : class
        {
            return GetExporterAttribute(typeof(T));
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
            return await ExportPdf(exporterAttribute.ToPdfExportOptions(), htmlString);
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
            return await ExportPdf(exporterAttribute.ToPdfExportOptions(), htmlString);
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
            return await ExportPdf(exporterAttribute.ToPdfExportOptions(), htmlString);
        }

        public async Task<byte[]> ExportListBytesByTemplate<T>(ICollection<T> data, PdfExportOptions pdfExportOptions, string template) where T : class
        {
            var htmlString = await HtmlExporter.ExportListByTemplate(data, template);
            return await ExportPdf(pdfExportOptions, htmlString);
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
            return await ExportPdf(pdfExporterAttribute.ToPdfExportOptions(), htmlString);
        }

        public async Task<byte[]> ExportBytesByTemplate<T>(T data, PdfExportOptions pdfExportOptions, string template) where T : class
        {
            var htmlString = await HtmlExporter.ExportByTemplate(data, template);
            return await ExportPdf(pdfExportOptions, htmlString);
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
            return await ExportPdf(pdfExporterAttribute.ToPdfExportOptions(), htmlString);
        }
    }
}
