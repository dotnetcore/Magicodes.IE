using System;
using System.Collections.Generic;
using System.Text;

namespace Magicodes.IE.Core
{
    /// <summary>
    /// 值映射
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
    public abstract class ValueMappingsBaseAttribute : Attribute
    {
        /// <summary>
        /// 根据字段类型获取映射
        /// </summary>
        /// <param name="fieldType"></param>
        /// <returns></returns>
        public abstract Dictionary<string, dynamic> GetMappings(Type fieldType);
    }
}
