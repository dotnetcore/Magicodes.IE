using OfficeOpenXml;

namespace Magicodes.JustCode.Import
{
    public class ColValue
    {
        public int Index { get; set; }
        public string Title { get; set; }
        public string Text { get; set; }
        public object Value { get; set; }
        public ExcelRange ExcelRange { get; set; }
    }
}
