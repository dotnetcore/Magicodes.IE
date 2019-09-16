using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace Magicodes.ExporterAndImporter.Excel.Utility
{
    /// <summary>
    /// 枚举辅助类
    /// </summary>
    public static class EnumHelper
    {
        /// <summary>
        /// 获取枚举的显示名称
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static IDictionary<string, int> GetDisplayNames(Type type)
        {
            if (!type.IsEnum)
            {
                throw new InvalidOperationException("无效的类型，请检查是否为枚举类型");
            }

            var names = Enum.GetNames(type);
            IDictionary<string, int> displayNames = new Dictionary<string, int>();
            foreach (var name in names)
            {
                var displayAttribute = type.GetField(name)
                    .GetCustomAttributes(typeof(DisplayAttribute), false)
                    .SingleOrDefault() as DisplayAttribute;
                if (displayAttribute != null)
                {
                    var value = (int)Enum.Parse(type, name);
                    displayNames.Add(displayAttribute.Name, value);
                }
            }

            return displayNames;
        }
    }
}