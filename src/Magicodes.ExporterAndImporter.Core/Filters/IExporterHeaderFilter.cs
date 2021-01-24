using Magicodes.ExporterAndImporter.Core.Models;

namespace Magicodes.ExporterAndImporter.Core.Filters
{
    /// <summary>
    /// 列头过滤
    /// </summary>
    public interface IExporterHeaderFilter : IFilter
    {
        /// <summary>
        /// 过滤列头（可以在此处理列名、是否隐藏等）
        /// </summary>
        /// <returns></returns>
        ExporterHeaderInfo Filter(ExporterHeaderInfo exporterHeaderInfo);
    }
}
