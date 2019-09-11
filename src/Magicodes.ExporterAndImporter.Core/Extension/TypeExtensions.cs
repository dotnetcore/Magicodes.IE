namespace Magicodes.ExporterAndImporter.Core.Extension
{
    using Microsoft.CSharp;
    using System;
    using System.CodeDom;
    using System.Collections;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.ComponentModel.DataAnnotations;
    using System.Linq;
    using System.Reflection;
    using System.Text;

    /// <summary>
    /// Defines the <see cref="TypeExtensions" />
    /// </summary>
    public static class TypeExtensions
    {
        
        /// <summary>
        /// 获取显示名
        /// </summary>
        /// <param name="customAttributeProvider"></param>
        /// <param name="inherit"></param>
        /// <returns></returns>
        public static string GetDisplayName(this ICustomAttributeProvider customAttributeProvider, bool inherit = false)
        {
            string displayName = null;
            var displayAttribute = customAttributeProvider.GetAttribute<DisplayAttribute>(false);
            if (displayAttribute != null) displayName = displayAttribute.Name;
            else
            {
                var displayNameAttribute = customAttributeProvider.GetAttribute<DisplayNameAttribute>(false);
                if (displayNameAttribute != null)
                    displayName = displayNameAttribute.DisplayName;
            }
            return displayName;
        }

        /// <summary>
        /// 获取类型描述
        /// </summary>
        /// <param name="customAttributeProvider"></param>
        /// <param name="inherit"></param>
        /// <returns></returns>
        public static string GetDescription(this ICustomAttributeProvider customAttributeProvider, bool inherit = false)
        {
            string des = string.Empty;
            var desAttribute = customAttributeProvider.GetAttribute<DescriptionAttribute>(false);
            if (desAttribute != null) des = desAttribute.Description;
            return des;
        }

        /// <summary>
        /// 获取类型描述或显示名
        /// </summary>
        /// <param name="customAttributeProvider"></param>
        /// <param name="inherit"></param>
        /// <returns></returns>
        public static string GetTypeDisplayOrDescription(this ICustomAttributeProvider customAttributeProvider, bool inherit = false)
        {
            var dispaly = customAttributeProvider.GetDescription(inherit);
            if (dispaly.IsNullOrWhiteSpace()) dispaly = customAttributeProvider.GetDisplayName(inherit);
            return dispaly ?? string.Empty;
        }


        /// <summary>
        /// 获取程序集属性
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="assembly"></param>
        /// <param name="inherit"></param>
        /// <returns></returns>
        public static T GetAttribute<T>(this ICustomAttributeProvider assembly, bool inherit = false)
 where T : Attribute
        {

            return assembly
                .GetCustomAttributes(typeof(T), inherit)
                .OfType<T>()
                .FirstOrDefault();
        }

        /// <summary>
        /// 检查指定指定类型成员中是否存在指定的Attribute特性
        /// </summary>
        /// <typeparam name="T">要检查的Attribute特性类型</typeparam>
        /// <param name="assembly">The assembly<see cref="ICustomAttributeProvider"/></param>
        /// <param name="inherit">是否从继承中查找</param>
        /// <returns>是否存在</returns>
        public static bool AttributeExists<T>(this ICustomAttributeProvider assembly, bool inherit = false) where T : Attribute
        {
            return assembly.GetCustomAttributes(typeof(T), inherit).Any(m => (m as T) != null);
        }


        /// <summary>
        /// 获取当前程序集中应用此特性的类
        /// </summary>
        /// <typeparam name="TAttribute"></typeparam>
        /// <param name="assembly"></param>
        /// <param name="inherit">The inherit<see cref="bool"/></param>
        /// <returns></returns>
        public static IEnumerable<Type> GetTypesWith<TAttribute>(this Assembly assembly, bool inherit) where TAttribute : System.Attribute
        {
            var attrType = typeof(TAttribute);
            foreach (Type type in assembly.GetTypes())
            {
                if (type.GetCustomAttributes(attrType, true).Length > 0)
                {
                    yield return type;
                }
            }
        }

        /// <summary>
        /// 获取枚举定义列表
        /// </summary>
        /// <typeparam name="TAttribute">枚举类型</typeparam>
        /// <returns>返回枚举列表元组（名称、值、描述）</returns>
        public static List<Tuple<string, int, string>> GetEnumDefinitionList(this Type type)
        {
            var list = new List<Tuple<string, int, string>>();
            var attrType = type;
            if (!attrType.IsEnum) return null;
            var names = Enum.GetNames(attrType);
            var values = Enum.GetValues(attrType);
            var index = 0;
            foreach (var value in values)
            {
                var name = names[index];
                string des = null;
                var objAttrs = value.GetType().GetField(value.ToString()).GetCustomAttributes(typeof(DescriptionAttribute), true);
                if (objAttrs != null &&
                    objAttrs.Length > 0)
                {
                    var descAttr = objAttrs[0] as DescriptionAttribute;
                    des = descAttr?.Description;
                }
                var item = new Tuple<string, int, string>(name, Convert.ToInt32(value), des);
                list.Add(item);
                index++;
            }
            return list;
        }

        /// <summary>
        /// 获取类的所有枚举
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static Dictionary<string, List<Tuple<string, int, string>>> GetClassEnumDefinitionList(this Type type)
        {
            var enumPros = type.GetProperties().Where(p => p.PropertyType.IsEnum);
            var dic = new Dictionary<string, List<Tuple<string, int, string>>>();
            foreach (var item in enumPros)
            {
                dic.Add(item.Name, item.PropertyType.GetEnumDefinitionList());
            }
            return dic;
        }
    }
}
