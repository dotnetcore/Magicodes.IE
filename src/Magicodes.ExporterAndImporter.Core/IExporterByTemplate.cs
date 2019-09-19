using Magicodes.ExporterAndImporter.Core.Models;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Core
{
    public interface IExporterByTemplate
    {
        /// <summary>
        /// 根据模板导出
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="fileName"></param>
        /// <param name="dataItems"></param>
        /// <param name="htmlTemplate">Html模板内容</param>
        /// <returns></returns>
        Task<TemplateFileInfo> ExportByTemplate<T>(string fileName, IList<T> dataItems, string htmlTemplate = null) where T : class;
    }
}
