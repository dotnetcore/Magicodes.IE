using Magicodes.ExporterAndImporter.Core;

namespace Magicodes.ExporterAndImporter.Excel
{
    /// <summary>
    /// Excel导出设计
    /// </summary>
    public class ExcelExporterAttribute : ExporterAttribute
    {
        /// <summary>
        /// 表格样式风格
        /// </summary>
        public string TableStyle { get; set; } = "Medium10";

        /// <summary>
        /// 自适应所有列
        /// </summary>
        public bool AutoFitAllColumn { get; set; }

        /// <summary>
        /// 作者
        /// </summary>
        public string Author { get; set; }
    }
}