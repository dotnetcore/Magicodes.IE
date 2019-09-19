using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Models;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Word
{
    /// <summary>
    /// Word导出
    /// </summary>
    public class WordExporter : IExporter, IExporterByTemplate
    {
        public Task<TemplateFileInfo> Export<T>(string fileName, IList<T> dataItems) where T : class => throw new NotImplementedException();
        public Task<byte[]> ExportAsByteArray<T>(IList<T> dataItems) where T : class => throw new NotImplementedException();
        public Task<TemplateFileInfo> ExportByTemplate<T>(string fileName, IList<T> dataItems, string htmlTemplate) where T : class
        {

        }
        public Task<byte[]> ExportHeaderAsByteArray(string[] items, string sheetName, ExcelHeadStyle globalStyle = null, List<ExcelHeadStyle> styles = null) => throw new NotImplementedException();
        public Task<byte[]> ExportHeaderAsByteArray<T>(T type) where T : class => throw new NotImplementedException();
    }
}
