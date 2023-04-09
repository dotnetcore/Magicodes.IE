using OfficeOpenXml;

namespace Magicodes.IE.Stash.Import
{
    /// <summary>
    /// 单元格
    /// </summary>
    public class CellValue
    {
        public int Index { get; set; }
        public string Title { get; set; }
        public string Text { get; set; }
        public object Value { get; set; }
        public ExcelRange ExcelRange { get; set; }
    }
}
