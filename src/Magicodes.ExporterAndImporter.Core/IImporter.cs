using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using Magicodes.ExporterAndImporter.Core.Models;

namespace Magicodes.ExporterAndImporter.Core
{
    /// <summary>
    ///     导入
    /// </summary>
    public interface IImporter
    {
        /// <summary>
        ///     导入为DataTable
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns>DataTable</returns>
        Task<DataTable> Import(string filePath);

        /// <summary>
        /// 导入为集合
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <returns></returns>
        Task<IList<T>> Import<T>(string filePath) where T : class, new();
    }
}