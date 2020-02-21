using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using Magicodes.ExporterAndImporter.Core.Filters;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Csv.Utility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Csv
{
    /// <summary>
    ///     Csv导出程序
    /// </summary>
    public class CsvExporter : IExporter
    {
        /// <summary>
        ///     导出Csv
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName">文件名</param>
        /// <param name="dataItems">数据列</param>
        /// <returns>文件</returns>
        public async Task<ExportFileInfo> Export<T>(string fileName, ICollection<T> dataItems) where T : class
        {
            fileName.CheckCsvFileName();
            var bytes = await ExportAsByteArray(dataItems);
            return bytes.ToCsvExportFileInfo(fileName);
        }
        /// <summary>
        ///     导出Csv
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <param name="dataItems"></param>
        /// <returns></returns>
        public async Task<ExportFileInfo> Export<T>(string fileName, DataTable dataItems) where T : class
        {
            fileName.CheckCsvFileName();
            var bytes = await ExportAsByteArray<T>(dataItems);
            return bytes.ToCsvExportFileInfo(fileName);
        }

        public Task<ExportFileInfo> Export(string fileName, DataTable dataItems, IExporterHeaderFilter exporterHeaderFilter = null, int maxRowNumberOnASheet = 1000000)
        {
            throw new NotImplementedException();
        }
        /// <summary>
        ///     导出字节
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataItems">数据列</param>
        /// <returns></returns>
        public Task<byte[]> ExportAsByteArray<T>(ICollection<T> dataItems) where T : class
        {
            var helper = new ExportHelper<T>();
            return Task.FromResult(helper.GetCsvExportAsByteArray(dataItems));
            
        }
        /// <summary>
        /// 导出DataTable
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataItems"></param>
        /// <returns></returns>
        public Task<byte[]> ExportAsByteArray<T>(DataTable dataItems) where T : class
        {
            var helper = new ExportHelper<T>();
            return Task.FromResult(helper.GetCsvExportAsByteArray<T>(dataItems));
        }

        public Task<byte[]> ExportAsByteArray(DataTable dataItems, IExporterHeaderFilter exporterHeaderFilter = null, int maxRowNumberOnASheet = 1000000)
        {
            throw new NotImplementedException();
        }

        public Task<byte[]> ExportHeaderAsByteArray(string[] items, string sheetName = "导出结果")
        {
            throw new NotImplementedException();
        }
        /// <summary>
        ///     导出Csv表头
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="type"></param>
        /// <returns></returns>
        public Task<byte[]> ExportHeaderAsByteArray<T>(T type) where T : class
        {
            var helper = new ExportHelper<T>();
            return Task.FromResult(helper.GetCsvExportAsByteArray());
        }
    }
}
