using System;

namespace Magicodes.ExporterAndImporter.Core
{
    public class ExporterCellAttribute : Attribute
    {
        public ExporterCellAttribute() {
        }
        public string DisplayName { set; get; }
        public float? FontSize { set; get; }
        public bool IsBold { set; get; }
        public string Format { get; set; }
    }
}