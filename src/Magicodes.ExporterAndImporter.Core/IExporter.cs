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

using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core.Filters;
using Magicodes.ExporterAndImporter.Core.Models;

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
        /// </summary>
        /// <param name="dataItems">数据</param>
        /// <returns>文件二进制数组</returns>
        Task<byte[]> ExportAsByteArray<T>(ICollection<T> dataItems) where T : class;

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
        ///     导出
        /// </summary>
        /// <param name="fileName">文件名称</param>
        /// <param name="dataItems">数据</param>
        /// <param name="exporterHeaderFilter">表头筛选器</param>
        /// <param name="maxRowNumberOnASheet">一个Sheet最大允许的行数，设置了之后将输出多个Sheet</param>
        /// <returns>文件</returns>
        Task<ExportFileInfo> Export(string fileName, DataTable dataItems, IExporterHeaderFilter exporterHeaderFilter = null, int maxRowNumberOnASheet = 1000000);

        /// <summary>
        ///     导出
        /// </summary>
        /// <param name="dataItems">数据</param>
        /// <param name="exporterHeaderFilter">表头筛选器</param>
        /// <param name="maxRowNumberOnASheet">一个Sheet最大允许的行数，设置了之后将输出多个Sheet</param>
        /// <returns>文件二进制数组</returns>
        Task<byte[]> ExportAsByteArray(DataTable dataItems, IExporterHeaderFilter exporterHeaderFilter = null, int maxRowNumberOnASheet = 1000000);

        /// <summary>
        ///     导出表头
        /// </summary>
        /// <param name="items">表头数组</param>
        /// <param name="sheetName">工作簿名称</param>
        /// <returns>文件二进制数组</returns>
        Task<byte[]> ExportHeaderAsByteArray(string[] items, string sheetName = "导出结果");

        /// <summary>
        ///     导出表头
        /// </summary>
        /// <param name="type">类型</param>
        /// <returns>文件二进制数组</returns>
        Task<byte[]> ExportHeaderAsByteArray<T>(T type) where T : class;
    }
}