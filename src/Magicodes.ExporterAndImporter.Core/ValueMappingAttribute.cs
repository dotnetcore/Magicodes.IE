// ======================================================================
// 
//           filename : ValueMappingAttribute.cs
//           description :
// 
//           created by 雪雁 at  2019-11-04 20:52
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using System;

namespace Magicodes.ExporterAndImporter.Core
{
    /// <summary>
    ///     值映射
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
    public class ValueMappingAttribute : Attribute
    {
        /// <summary>
        ///     设置文本和值映射
        /// </summary>
        /// <param name="text">文本</param>
        /// <param name="value">值</param>
        public ValueMappingAttribute(string text, object value)
        {
            Text = text;
            Value = value;
        }

        /// <summary>
        ///     文本
        /// </summary>
        public string Text { get; }

        /// <summary>
        ///     值
        /// </summary>
        public object Value { get; }
    }
}