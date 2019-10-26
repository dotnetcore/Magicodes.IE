// ======================================================================
// 
//           Copyright (C) 2019-2030 湖南心莱信息科技有限公司
//           All rights reserved
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

using DinkToPdf;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Html;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Pdf
{
    /// <summary>
    /// Pdf导出逻辑
    /// </summary>
    public class PdfExporter : IExporterByTemplate
    {
        /// <summary>
        ///     根据模板导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataItems"></param>
        /// <param name="htmlTemplate">Html模板内容</param>
        /// <returns></returns>
        public Task<string> ExportListByTemplate<T>(ICollection<T> dataItems, string htmlTemplate = null) where T : class => throw new NotImplementedException();

        /// <summary>
        ///     根据模板导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="htmlTemplate">Html模板内容</param>
        /// <returns></returns>
        public Task<string> ExportByTemplate<T>(T data, string htmlTemplate = null) where T : class
        {
            throw new NotImplementedException();
        }

        /// <summary>
        ///    根据模板导出列表
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <param name="dataItems"></param>
        /// <param name="htmlTemplate"></param>
        /// <returns></returns>
        public async Task<TemplateFileInfo> ExportListByTemplate<T>(string fileName, ICollection<T> dataItems, string htmlTemplate = null) where T : class
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new ArgumentException("文件名必须填写!", nameof(fileName));
            }

            var exporter = new HtmlExporter();
            var htmlString = await exporter.ExportListByTemplate(dataItems, htmlTemplate);
            var converter = new BasicConverter(new PdfTools());
            var doc = new HtmlToPdfDocument
            {
                GlobalSettings =
                {
                    ColorMode = ColorMode.Color,
                    Orientation = Orientation.Landscape,
                    PaperSize = PaperKind.A4,
                    Out = fileName,
                },
                Objects =
                {
                    new ObjectSettings
                    {
                        PagesCount = true,
                        HtmlContent = htmlString,
                        WebSettings = {DefaultEncoding = "utf-8"},
                        HeaderSettings = {FontSize = 9, Right = "[page]/[toPage]", Line = true, Spacing = 2.812},
                        Encoding = System.Text.Encoding.UTF8
                    }
                }
            };
            converter.Convert(doc);
            var fileInfo = new TemplateFileInfo(fileName, "application/pdf");
            return fileInfo;
        }

        /// <summary>
        ///     根据模板导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <param name="data"></param>
        /// <param name="htmlTemplate"></param>
        /// <returns></returns>
        public async Task<TemplateFileInfo> ExportByTemplate<T>(string fileName, T data, string htmlTemplate) where T : class
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                throw new ArgumentException("文件名必须填写!", nameof(fileName));
            }

            var exporter = new HtmlExporter();
            var htmlString = await exporter.ExportByTemplate(data, htmlTemplate);
            var converter = new BasicConverter(new PdfTools());
            var doc = new HtmlToPdfDocument
            {
                GlobalSettings =
                {
                    ColorMode = ColorMode.Color,
                    Orientation = Orientation.Landscape,
                    PaperSize = PaperKind.A4,
                    Out = fileName,
                    
                },
                Objects =
                {
                    new ObjectSettings
                    {
                        //PagesCount = true,
                        HtmlContent = htmlString,
                        WebSettings = {DefaultEncoding = "utf-8"},
                        Encoding = System.Text.Encoding.UTF8
                        //HeaderSettings = {FontSize = 9, Right = "[page]/[toPage]", Line = true, Spacing = 2.812},
                    }
                }
            };
            converter.Convert(doc);
            var fileInfo = new TemplateFileInfo(fileName, "application/pdf");
            return fileInfo;
        }
    }
}