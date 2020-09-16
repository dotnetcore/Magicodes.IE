using Magicodes.ExporterAndImporter.Core.Models;

namespace Magicodes.ExporterAndImporter.Core.Filters
{
    /// <summary>
    /// 导入结果筛选器
    /// 可以处理标注内容
    /// </summary>
    public interface IImportResultFilter : IFilter
    {
        /// <summary>
        /// 处理导入结果
        /// 比如对错误信息进行多语言转换
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        ImportResult<T> Filter<T>(ImportResult<T> importResult) where T : class, new();
    }
}
