using CsvHelper;
using Magicodes.ExporterAndImporter.Core.Models;
using System;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Csv.Utility
{
    /// <summary>
    ///     导入辅助类
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ImportHelper<T> : IDisposable where T : class, new()
    {
        /// <summary>
        /// </summary>
        /// <param name="filePath"></param>
        public ImportHelper(string filePath = null)
        {
            FilePath = filePath;
        }
        /// <summary>
        ///     导入文件路径
        /// </summary>
        protected string FilePath { get; set; }

        /// <summary>
        ///     导入结果
        /// </summary>
        internal ImportResult<T> ImportResult { get; set; }

        /// <summary>
        ///     导入模型
        /// </summary>
        /// <returns></returns>
        public Task<ImportResult<T>> Import(string filePath = null)
        {
            if (!string.IsNullOrWhiteSpace(filePath)) FilePath = filePath;

            ImportResult = new ImportResult<T>();
            try
            {
                CheckImportFile(FilePath);

                using (var reader = new System.IO.StreamReader(FilePath))
                using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
                {
                    csv.Configuration.RegisterClassMap<AutoMap<T>>();
                    var result = csv.GetRecords<T>();
                    ImportResult.Data = result.ToList();
                    return Task.FromResult(ImportResult);
                }
            }
            catch (Exception ex)
            {
                ImportResult.Exception = ex;
            }
            return Task.FromResult(ImportResult);
        }
        /// <summary>
        ///     检查导入文件路径
        /// </summary>
        /// <exception cref="ArgumentException">文件路径不能为空! - filePath</exception>
        private static void CheckImportFile(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("文件路径不能为空!", nameof(filePath));

            //TODO:在Docker容器中存在文件路径找不到问题，暂时先注释掉
            //if (!File.Exists(filePath))
            //{
            //    throw new ImportException("导入文件不存在!");
            //}
        }
        /// <summary>
        /// </summary>
        public void Dispose()
        {
            FilePath = null;
            ImportResult = null;
            GC.Collect();
        }


    }
}
