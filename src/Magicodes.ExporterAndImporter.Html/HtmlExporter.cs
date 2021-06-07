// ======================================================================
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

using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Core.Models;
using RazorEngine;
using RazorEngine.Configuration;
using RazorEngine.Templating;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Encoding = System.Text.Encoding;

namespace Magicodes.ExporterAndImporter.Html
{
    /// <summary>
    ///     HTML导出
    /// </summary>
    public partial class HtmlExporter : IExporterByTemplate
    {
        /// <summary>
        /// 初始化
        /// </summary>
        public HtmlExporter()
        {
            //配置RazorEngine
            var config = new TemplateServiceConfiguration()
            {
                ReferenceResolver = new ExternalAssemblyReferenceResolver(null)
            };
            Engine.Razor = RazorEngineService.Create(config);
        }

        /// <summary>
        ///     根据模板导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataItems"></param>
        /// <param name="htmlTemplate">Html模板内容</param>
        /// <returns></returns>
        public Task<string> ExportListByTemplate<T>(ICollection<T> dataItems, string htmlTemplate = null)
            where T : class
        {
            var result = RunCompileTpl(new ExportDocumentInfoOfListData<T>(dataItems), htmlTemplate);
            return Task.FromResult(result);
        }

        /// <summary>
        ///     根据模板导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="htmlTemplate">Html模板内容</param>
        /// <returns></returns>
        public Task<string> ExportByTemplate<T>(T data, string htmlTemplate) where T : class
        {
            var result = RunCompileTpl(new ExportDocumentInfo<T>(data), htmlTemplate);
            return Task.FromResult(result);
        }
        /// <summary>
        ///     根据模板导出
        /// </summary>
        /// <param name="data"></param>
        /// <param name="htmlTemplate">Html模板内容</param>
        /// <param name="type"></param>
        /// <returns></returns>
        public Task<string> ExportByTemplate(object data, string htmlTemplate, Type type)
        {
            var result = RunCompileTpl(new ExportDocumentInfo(data, type), type, htmlTemplate);
            return Task.FromResult(result);
        }
        /// <summary>
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <param name="dataItems"></param>
        /// <param name="htmlTemplate"></param>
        /// <returns></returns>
        public async Task<ExportFileInfo> ExportListByTemplate<T>(string fileName, ICollection<T> dataItems,
            string htmlTemplate = null) where T : class
        {
            var file = new ExportFileInfo(fileName, "text/html");

            var result = await ExportListByTemplate(dataItems, htmlTemplate);
            File.WriteAllText(fileName, result, Encoding.UTF8);
            return file;
        }

        /// <summary>
        /// 导出HTML文件
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <param name="data"></param>
        /// <param name="htmlTemplate"></param>
        /// <returns></returns>
        public async Task<ExportFileInfo> ExportByTemplate<T>(string fileName, T data,
            string htmlTemplate) where T : class
        {
            var file = new ExportFileInfo(fileName, "text/html");
            var result = await ExportByTemplate(data, htmlTemplate);

            File.WriteAllText(fileName, result, Encoding.UTF8);
            return file;
        }

        /// <summary>
        ///     获取HTML模板
        /// </summary>
        /// <param name="htmlTemplate"></param>
        /// <returns></returns>
        protected string GetHtmlTemplate(string htmlTemplate = null)
        {
            return string.IsNullOrWhiteSpace(htmlTemplate)
                ? typeof(HtmlExporter).Assembly.ReadManifestString("default.cshtml")
                : htmlTemplate;
        }


        /// <summary>
        ///     编译和运行模板
        /// </summary>
        /// <param name="model"></param>
        /// <param name="htmlTemplate"></param>
        /// <returns></returns>
        protected string RunCompileTpl(object model, string htmlTemplate = null)
        {
            var htmlTpl = GetHtmlTemplate(htmlTemplate);
            return Engine.Razor.RunCompile(htmlTpl, htmlTpl.GetHashCode().ToString(), null, model);
        }

        /// <summary>
        ///     编译和运行模板
        /// </summary>
        /// <param name="model"></param>
        /// <param name="type"></param>
        /// <param name="htmlTemplate"></param>
        /// <returns></returns>
        protected string RunCompileTpl(object model, Type type, string htmlTemplate = null)
        {
            var htmlTpl = GetHtmlTemplate(htmlTemplate);

            return Engine.Razor.RunCompile(htmlTpl, htmlTpl.GetHashCode().ToString(), null, model);
        }

        /// <summary>
        /// 导出到bytes
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="template"></param>
        /// <returns></returns>
        public async Task<byte[]> ExportListBytesByTemplate<T>(ICollection<T> data,
            string template) where T : class
        {
            var result = await ExportListByTemplate(data, template);
            return Encoding.UTF8.GetBytes(result);
        }
        /// <summary>
        ///  导出到bytes
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="template"></param>
        /// <returns></returns>
        public async Task<byte[]> ExportBytesByTemplate<T>(T data,
            string template) where T : class
        {
            var result = await ExportByTemplate(data, template);
            return Encoding.UTF8.GetBytes(result);
        }
        /// <summary>
        ///     
        /// </summary>
        /// <param name="data"></param>
        /// <param name="template"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        public async Task<byte[]> ExportBytesByTemplate(object data, string template, Type type)
        {
            var result = await ExportByTemplate(data, template, type);
            return Encoding.UTF8.GetBytes(result);
        }
    }
}