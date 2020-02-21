using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Models;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Csv.Utility;

namespace Magicodes.ExporterAndImporter.Csv
{
    /// <summary>
    ///     Csv导入
    /// </summary>
    public class CsvImporter : IImporter
    {
        public Task<ExportFileInfo> GenerateTemplate<T>(string fileName) where T : class, new()
        {
            throw new NotImplementedException();
        }

        public Task<byte[]> GenerateTemplateBytes<T>() where T : class, new()
        {
            throw new NotImplementedException();
        }
        /// <summary>
        ///     导入
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <param name="labelingFilePath"></param>
        /// <returns></returns>
        public Task<ImportResult<T>> Import<T>(string filePath, string labelingFilePath = null) where T : class, new()
        {
            using (var importer = new ImportHelper<T>(filePath))
            {
                return importer.Import();
            }
        }

        public Task<Dictionary<string, ImportResult<object>>> ImportMultipleSheet<T>(string filePath) where T : class, new()
        {
            throw new NotImplementedException();
        }

        public Task<Dictionary<string, ImportResult<TSheet>>> ImportSameSheets<T, TSheet>(string filePath)
            where T : class, new()
            where TSheet : class, new()
        {
            throw new NotImplementedException();
        }
    }
}
