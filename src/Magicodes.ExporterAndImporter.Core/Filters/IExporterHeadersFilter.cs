using Magicodes.ExporterAndImporter.Core.Models;
using System.Collections.Generic;

namespace Magicodes.ExporterAndImporter.Core.Filters
{
    /// <summary>
    /// 列头（集合）过滤
    /// </summary>
    public interface IExporterHeadersFilter : IFilter
    {
        /// <summary>
        /// 过滤列头（集合）（可以在此处理列名、是否隐藏等）
        /// </summary>
        /// <returns></returns>
        IList<ExporterHeaderInfo> Filter(IList<ExporterHeaderInfo> exporterHeaderInfos);
    }
}
