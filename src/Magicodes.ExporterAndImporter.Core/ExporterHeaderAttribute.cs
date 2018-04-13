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

        /// <summary>
        /// 显示名称
        /// </summary>
        public string DisplayName { set; get; }

        /// <summary>
        /// 字体大小
        /// </summary>
        public float? FontSize { set; get; }

        /// <summary>
        /// 是否加粗
        /// </summary>
        public bool IsBold { set; get; }

        /// <summary>
        /// 格式化
        /// </summary>
        public string Format { get; set; }

        /// <summary>
        /// 是否自适应
        /// </summary>
        public bool IsAutoFit { set; get; }

        /// <summary>
        /// 是否忽略
        /// </summary>
        public bool IsIgnore { get; set; }
    }
}