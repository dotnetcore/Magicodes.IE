using System;
using System.Collections.Generic;
using System.Text;

namespace Magicodes.ExporterAndImporter.Excel
{
    /// <summary>
    /// 值映射
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
    public class ValueMappingAttribute : Attribute
    {
        /// <summary>
        /// 文本
        /// </summary>
        public string Text { get; }

        /// <summary>
        /// 值
        /// </summary>
        public object Value { get; }

        /// <summary>
        /// 设置文本和值映射
        /// </summary>
        /// <param name="text">文本</param>
        /// <param name="value">值</param>
        public ValueMappingAttribute(string text, object value)
        {
            Text = text;
            Value = value;
        }
    }
}
