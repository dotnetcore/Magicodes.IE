using Magicodes.ExporterAndImporter.Core;

namespace Magicodes.ExporterAndImporter.Excel
{
    /// <summary>
    /// Excel导入配置
    /// </summary>
    public class ExcelImporterAttribute : ImporterAttribute
    {
        /// <summary>
        /// 指定Sheet名称(获取指定Sheet名称)
        /// 为空则自动获取第一个
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// 最大导入行数
        /// 默认值为65000
        /// </summary>
        public int MaxRowNumber { get; set; } = 65000;
    }
}