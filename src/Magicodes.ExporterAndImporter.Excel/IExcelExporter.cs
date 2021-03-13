using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Filters;
using Magicodes.ExporterAndImporter.Core.Models;
using System.Collections.Generic;
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
        Task<ExportFileInfo> Export(string fileName, DataTable dataItems,
            IExporterHeaderFilter exporterHeaderFilter = null, int maxRowNumberOnASheet = 1000000);

        /// <summary>
        ///     导出
        /// </summary>
        /// <param name="dataItems">数据</param>
        /// <param name="exporterHeaderFilter">表头筛选器</param>
        /// <param name="maxRowNumberOnASheet">一个Sheet最大允许的行数，设置了之后将输出多个Sheet</param>
        /// <returns>文件二进制数组</returns>
        Task<byte[]> ExportAsByteArray(DataTable dataItems, IExporterHeaderFilter exporterHeaderFilter = null,
            int maxRowNumberOnASheet = 1000000);


        /// <summary>
        ///     追加集合到当前导出程序
        ///     append the collection to context
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataItems"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        ExcelExporter Append<T>(ICollection<T> dataItems, string sheetName = null) where T : class, new();

        /// <summary>
        ///    分割sheet追加当前column 
        /// </summary>
        /// <returns></returns>
        ExcelExporter SeparateByColumn();

        /// <summary>
        ///     分割导出多个sheet
        /// </summary>
        /// <returns></returns>
        ExcelExporter SeparateBySheet();

        /// <summary>
        ///     将rows追加到当前sheet
        /// </summary>
        /// <returns></returns>
        ExcelExporter SeparateByRow();

        /// <summary>
        ///     导出所有的追加数据
        ///     export excel after append all collectioins
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        Task<ExportFileInfo> ExportAppendData(string fileName);

        /// <summary>
        ///     导出所有的追加数据
        ///     export excel after append all collectioins
        /// </summary>
        /// <returns></returns>
        Task<byte[]> ExportAppendDataAsByteArray();
    }
}