using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Models;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Encoding = System.Text.Encoding;

namespace Magicodes.ExporterAndImporter.Html
{
    /// <summary>
    /// HTML导出
    /// </summary>
    public class HtmlExporter : IExporterByTemplate
    {
        /// <summary>
        /// 根据模板导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <param name="dataItems"></param>
        /// <param name="htmlTemplate">Html模板内容</param>
        /// <returns></returns>
        public async Task<TemplateFileInfo> ExportByTemplate<T>(string fileName, IList<T> dataItems, string htmlTemplate = null) where T : class
        {
            if (string.IsNullOrWhiteSpace(htmlTemplate))
            {
                var defaultHtmlTpl = ReadManifestData<HtmlExporter>("default.html");

                var script = CSharpScript.Create<string>("string htmlStr=\"\";");
                var matches = Regex.Matches(defaultHtmlTpl, @"@foreach[\w\s{}.<>()/\[\]\+\-?|\\*`$]+}", RegexOptions.Multiline);
                foreach (Match match in matches)
                {
                    if (match.Success)
                    {
                        //var sb = new StringBuilder();
                        var list = Regex.Split(match.Value, Environment.NewLine).Where(p => p.Trim() != "{" && p.Trim() != "}")
                            .ToList();
                        foreach (var dataItem in dataItems)
                        {
                            foreach (var line in list)
                            {
                                script.ContinueWith("htmlStr +=\"\\n\";");
                                if (line.Contains("{"))
                                {
                                    var codeStr = await CSharpScript.RunAsync<string>("var res=$\"" + line + "\";", globals: dataItem);
                                    var codeRes = await codeStr.ContinueWithAsync("res");
                                    script.ContinueWith("htmlStr +=$\"" + codeRes.ReturnValue + "\";");
                                }
                                else
                                {
                                    script.ContinueWith("htmlStr +=$\"" + line + "\";");
                                }

                                //var res = Regex.Matches(line, @"{{[\w\s.()/\[\]\+\-?|\\*`$]+}}", RegexOptions.None);
                                //if (res.Count == 0)
                                //{
                                //    script.ContinueWith("htmlStr +=\"" + line + "\"");
                                //}
                                //else
                                //{
                                //    script.ContinueWith("htmlStr +=$\"" + line + "\"");
                                //}
                            }
                        }



                    }
                }
                var result = await script.RunAsync();
                var a = await result.ContinueWithAsync("res");
                var test = a.ReturnValue;
                //var result = script.;
                //var script = CSharpScript.Create<int>("X*Y", globalsType: typeof(IList<T>));
                //script.Compile();
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
                using (var reader = new StreamReader(stream, encoding: Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
        }
    }
}
