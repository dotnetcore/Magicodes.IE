using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Filters;
using Magicodes.ExporterAndImporter.Core.Models;
using System.Data;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Excel
{
    /// <summary>
    /// Excel导出程序
    /// </summary>
    public interface IExcelExporter : IExporter, IExportFileByTemplate
    {
        /// <summary>
        ///     导出表头
        /// </summary>
        /// <param name="items">表头数组</param>
        /// <param name="sheetName">工作簿名称</param>
        /// <returns>文件二进制数组</returns>
        Task<byte[]> ExportHeaderAsByteArray(string[] items, string sheetName = "导出结果");
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
    }
}
