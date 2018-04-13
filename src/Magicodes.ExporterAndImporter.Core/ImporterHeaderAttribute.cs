using System;

namespace Magicodes.ExcelImporter.Attributes
{
    /// <summary>
    /// 导入头部信息
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ImporterHeaderAttribute : Attribute
    {
        public ImporterHeaderAttribute( float fontSize = 12, string format = null, bool bold = true, bool autoFit = true , bool required = false)
        {
            FontSize = fontSize;
            Format = format;
            Bold = bold;
            AutoFit = autoFit;
            Required = required;
        }

        public float FontSize { set; get; }
        public bool Bold { set; get; }
        public string Format { get; set; }
        public bool AutoFit { set; get; }
        public bool Required { set; get; }
    }
}