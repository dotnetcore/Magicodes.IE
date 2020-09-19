// ======================================================================
//
//           filename : IExportFileByTemplate.cs
//           description :
//
//           created by 雪雁 at  --
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
//
// ======================================================================

using Magicodes.ExporterAndImporter.Core.Models;
using System;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Core
{
    /// <summary>
    /// 根据模板导出文件
    /// </summary>
    public interface IExportFileByTemplate
    {
        /// <summary>
        ///     根据模板导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <param name="data"></param>
        /// <param name="template">HTML模板或模板路径</param>
        /// <returns></returns>
        Task<ExportFileInfo> ExportByTemplate<T>(string fileName, T data,
            string template) where T : class;

        /// <summary>
        ///     根据模板导出到载荷
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <param name="template">HTML模板或模板路径</param>
        /// <returns></returns>
        Task<byte[]> ExportBytesByTemplate<T>(T data, string template) where T : class;
        /// <summary>
        ///		根据模板导出
        /// </summary>
        /// <param name="data"></param>
        /// <param name="template"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        Task<byte[]> ExportBytesByTemplate(object data, string template, Type type);
    }
}