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
        /// 名称(比如当前Sheet 名称)
        /// </summary>
        public string Name { get; set; }

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