using System;

namespace Magicodes.ExporterAndImporter.Core
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ExporterHeaderAttribute : Attribute
    {
        public ExporterHeaderAttribute(string displayName = null, float fontSize = 11, string format = null,
            bool isBold = true, bool isAutoFit = true)
        {
            DisplayName = displayName;
            FontSize = fontSize;
            Format = format;
            IsBold = isBold;
            IsAutoFit = isAutoFit;
        }

        public string DisplayName { set; get; }
        public float? FontSize { set; get; }
        public bool IsBold { set; get; }
        public string Format { get; set; }
        public bool IsAutoFit { set; get; }
    }
}