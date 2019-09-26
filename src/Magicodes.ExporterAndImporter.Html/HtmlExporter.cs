// ======================================================================
// 
//           Copyright (C) 2019-2030 湖南心莱信息科技有限公司
//           All rights reserved
// 
//           filename : HtmlExporter.cs
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
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Models;
using RazorEngine;
using RazorEngine.Templating;
using Encoding = System.Text.Encoding;

namespace Magicodes.ExporterAndImporter.Html
{
    /// <summary>
    ///     HTML导出
    /// </summary>
    public class HtmlExporter : IExporterByTemplate
    {
        /// <summary>
        ///     根据模板导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataItems"></param>
        /// <param name="htmlTemplate">Html模板内容</param>
        /// <returns></returns>
        public async Task<string> ExportByTemplate<T>(IList<T> dataItems, string htmlTemplate = null) where T : class
        {
            var htmlTpl = string.IsNullOrWhiteSpace(htmlTemplate)
                ? ReadManifestData<HtmlExporter>("default.cshtml")
                : htmlTemplate;

            var exportDocumentInfo = new ExportDocumentInfo<T>(dataItems);
            var result =
                Engine.Razor.RunCompile(htmlTpl, htmlTpl.GetHashCode().ToString(), null, exportDocumentInfo);
            return result;
        }

        /// <summary>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <param name="dataItems"></param>
        /// <param name="htmlTemplate"></param>
        /// <returns></returns>
        public async Task<TemplateFileInfo> ExportByTemplate<T>(string fileName, IList<T> dataItems,
            string htmlTemplate = null) where T : class
        {
            var file = new TemplateFileInfo(fileName, "text/html");
            var result = await ExportByTemplate(dataItems, htmlTemplate);
            File.WriteAllText(fileName, result, Encoding.UTF8);
            return file;
        }

        public static string ReadManifestData<TSource>(string embeddedFileName) where TSource : class
        {
            var assembly = typeof(TSource).GetTypeInfo().Assembly;
            var resourceName = assembly.GetManifestResourceNames().First(s =>
                s.EndsWith(embeddedFileName, StringComparison.CurrentCultureIgnoreCase));

            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null) throw new InvalidOperationException("Could not load manifest resource stream.");
                using (var reader = new StreamReader(stream, Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
        }
    }
}