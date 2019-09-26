using DinkToPdf;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Html;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Pdf
{
    public class PdfExporter : IExporterByTemplate
    {

        /// <summary>
        /// 根据HTML模板导出PDF
        /// </summary>
        /// <param name="dataItems"></param>
        /// <param name="fileName"></param>
        /// <param name="htmlTemplate"></param>
        /// <returns></returns>
        public async Task<TemplateFileInfo> ExportByTemplate<T>(string fileName, IList<T> dataItems, string htmlTemplate = null) where T : class
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new ArgumentException("文件名必须填写!", nameof(fileName));
            }
            var exporter = new HtmlExporter();
            var htmlString = await exporter.ExportByTemplate(dataItems, htmlTemplate);
            var converter = new BasicConverter(new PdfTools());
            var doc = new HtmlToPdfDocument()
            {
                GlobalSettings = {
                    ColorMode = ColorMode.Color,
                    Orientation = Orientation.Landscape,
                    PaperSize = PaperKind.A4Plus,
                    Out = fileName
                },
                Objects = {
                    new ObjectSettings() {
                        PagesCount = true,
                        HtmlContent = htmlString,
                        WebSettings = { DefaultEncoding = "utf-8" },
                        HeaderSettings = { FontSize = 9, Right = "Page [page] of [toPage]", Line = true, Spacing = 2.812 }
                    }
                }
            };
            converter.Convert(doc);
            var fileInfo = new TemplateFileInfo(fileName, "application/pdf");
            return fileInfo;
        }

        /// <summary>
        /// 根据模板导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataItems"></param>
        /// <param name="htmlTemplate">Html模板内容</param>
        /// <returns></returns>
        public Task<string> ExportByTemplate<T>(IList<T> dataItems, string htmlTemplate = null) where T : class => throw new NotImplementedException();
    }
}
