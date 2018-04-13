using System;

namespace Magicodes.ExporterAndImporter.Core
{
    [AttributeUsage(AttributeTargets.Class)]
    public class ExporterAttribute : Attribute
    {
        public ExporterAttribute()
        {
        }

        /// <summary>
        /// 头部字体大小
        /// </summary>
        public float? HeaderFontSize { set; get; }

        /// <summary>
        /// 正文字体大小
        /// </summary>
        public float? FontSize { set; get; }
    }
}