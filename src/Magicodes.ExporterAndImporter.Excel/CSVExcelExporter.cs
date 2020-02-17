using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Filters;
using Magicodes.ExporterAndImporter.Core.Models;

namespace Magicodes.ExporterAndImporter.Excel
{
    /// <summary>
    ///     CSV导出程序
    /// </summary>
    public class CSVExcelExporter 
    {
        public Task<ExportFileInfo> Export<T>(string fileName, ICollection<T> dataItems) where T : class
        {
            throw new NotImplementedException();
        }

        public Task<ExportFileInfo> Export<T>(string fileName, DataTable dataItems) where T : class
        {
            throw new NotImplementedException();
        }

        public Task<ExportFileInfo> Export(string fileName, DataTable dataItems, IExporterHeaderFilter exporterHeaderFilter = null, int maxRowNumberOnASheet = 1000000)
        {
            throw new NotImplementedException();
        }

        public Task<byte[]> ExportAsByteArray<T>(ICollection<T> dataItems) where T : class
        {
            throw new NotImplementedException();
        }

        public Task<byte[]> ExportAsByteArray<T>(DataTable dataItems) where T : class
        {
            throw new NotImplementedException();
        }

        public Task<byte[]> ExportAsByteArray(DataTable dataItems, IExporterHeaderFilter exporterHeaderFilter = null, int maxRowNumberOnASheet = 1000000)
        {
            throw new NotImplementedException();
        }

        public Task<byte[]> ExportHeaderAsByteArray(string[] items, string sheetName = "导出结果")
        {
            throw new NotImplementedException();
        }

        public Task<byte[]> ExportHeaderAsByteArray<T>(T type) where T : class
        {
            throw new NotImplementedException();
        }
    }
}
