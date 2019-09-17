using Magicodes.ExporterAndImporter.Core;

namespace Magicodes.ExporterAndImporter.Excel
{
    /// <summary>
    /// Excel导出表头信息
    /// </summary>
    public class ExporterHeaderInfo
    {
        /// <summary>
        /// 列索引
        /// </summary>
        public int Index { get; set; }

        /// <summary>
        /// 列名称
        /// </summary>
        public string PropertyName { get; set; }

        /// <summary>
        /// 列属性
        /// </summary>
        public ExporterHeaderAttribute ExporterHeader { get; set; }
    }
}