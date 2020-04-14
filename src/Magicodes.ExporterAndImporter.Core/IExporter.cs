// ======================================================================
// 
//           filename : IExporter.cs
//           description :
// 
//           created by 雪雁 at  2019-09-11 13:51
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using Magicodes.ExporterAndImporter.Core.Models;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Core
{
    /// <summary>
    ///     导出
    /// </summary>
    public interface IExporter
    {
        /// <summary>
        ///     导出
        /// </summary>
        /// <param name="fileName">文件名称</param>
        /// <param name="dataItems">数据</param>
        /// <returns>文件</returns>
        Task<ExportFileInfo> Export<T>(string fileName, ICollection<T> dataItems) where T : class;

        /// <summary>
        ///     导出
        ///     Support export through configuration
        ///     支持通过配置导出
        ///     https://github.com/dotnetcore/Magicodes.IE/issues/61
        /// </summary>
        /// <param name="dataItems">数据</param>
        /// <param name="exporterHeaders">导出表头设置</param>
        /// <returns>文件二进制数组</returns>
        Task<byte[]> ExportAsByteArray<T>(ICollection<T> dataItems, List<ExporterHeaderInfo> exporterHeaders = null) where T : class;

        /// <summary>
        ///     导出
        /// </summary>
        /// <param name="fileName">文件名称</param>
        /// <param name="dataItems">数据</param>
        /// <returns>文件</returns>
        Task<ExportFileInfo> Export<T>(string fileName, DataTable dataItems) where T : class;

        /// <summary>
        ///     导出
        /// </summary>
        /// <param name="dataItems">数据</param>
        /// <returns>文件二进制数组</returns>
        Task<byte[]> ExportAsByteArray<T>(DataTable dataItems) where T : class;

        /// <summary>
        ///     导出表头
        /// </summary>
        /// <param name="type">类型</param>
        /// <returns>文件二进制数组</returns>
        Task<byte[]> ExportHeaderAsByteArray<T>(T type) where T : class;
    }
}