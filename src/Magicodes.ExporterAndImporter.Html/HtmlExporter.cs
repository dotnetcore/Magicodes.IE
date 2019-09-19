using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Models;
using System.Reflection;
using RazorEngine;
using RazorEngine.Configuration;
using Encoding = System.Text.Encoding;
using RazorEngine.Templating;

namespace Magicodes.ExporterAndImporter.Html
{
    /// <summary>
    /// HTML导出
    /// </summary>
    public class HtmlExporter: IExporterByTemplate
    {
        /// <summary>
        /// 根据模板导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <param name="dataItems"></param>
        /// <param name="htmlTemplate">Html模板内容</param>
        /// <returns></returns>
        public Task<TemplateFileInfo> ExportByTemplate<T>(string fileName, IList<T> dataItems, string htmlTemplate = null) where T : class
        {
            if (string.IsNullOrWhiteSpace(htmlTemplate))
            {
                var defaultHtmlTpl = ReadManifestData<HtmlExporter>("default.cshtml");
                var config = new TemplateServiceConfiguration()
                {
                    
                };
                var service = RazorEngineService.Create(config);
                Engine.Razor = service;
                var res = Engine.Razor.RunCompile(defaultHtmlTpl, fileName, typeof(IList<T>), dataItems);
                var t = new TemplateFileInfo();
            }
            throw new NotImplementedException();
        }

        public static string ReadManifestData<TSource>(string embeddedFileName) where TSource : class
        {
            var assembly = typeof(TSource).GetTypeInfo().Assembly;
            var resourceName = assembly.GetManifestResourceNames().First(s => s.EndsWith(embeddedFileName, StringComparison.CurrentCultureIgnoreCase));

            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    throw new InvalidOperationException("Could not load manifest resource stream.");
                }
                using (var reader = new StreamReader(stream,encoding: Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
        }
    }
}
