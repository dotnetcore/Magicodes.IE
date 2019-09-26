using System.Collections.Generic;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Core
{
    /// <summary>
    /// 
    /// </summary>
    public interface IExporterByTemplate
    {
        /// <summary>
        /// 根据模板导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataItems"></param>
        /// <param name="htmlTemplate">Html模板内容</param>
        /// <returns></returns>
        Task<string> ExportByTemplate<T>(IList<T> dataItems, string htmlTemplate = null) where T : class;
    }
}
