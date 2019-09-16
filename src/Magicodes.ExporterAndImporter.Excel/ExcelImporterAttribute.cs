using Magicodes.ExporterAndImporter.Core;

namespace Magicodes.ExporterAndImporter.Excel
{
    public class ExcelImporterAttribute : ImporterAttribute
    {
        public ExcelImporterAttribute()
        {
        }

        /// <summary>
        /// 指定Sheet名称(获取指定Sheet名称)
        /// 为空则自动获取第一个
        /// </summary>
        public string SheetName { get; set; }
    }
}