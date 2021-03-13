using Magicodes.ExporterAndImporter.Core.Models;
using System.Collections.Generic;

namespace Magicodes.ExporterAndImporter.Core.Filters
{
    /// <summary>
    /// 导入列头筛选器
    /// 可以自行处理列头设置、值映射等
    /// </summary>
    public interface IImportHeaderFilter : IFilter
    {
        /// <summary>
        /// 处理列头
        /// </summary>
        /// <param name="importerHeaderInfos"></param>
        /// <returns></returns>
        List<ImporterHeaderInfo> Filter(List<ImporterHeaderInfo> importerHeaderInfos);
    }
}
