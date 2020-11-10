// ======================================================================
// 
//           filename : TypeExtensions.cs
//           description :
// 
//           created by 雪雁 at  2019-09-11 16:53
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using Magicodes.ExporterAndImporter.Core.Filters;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;
#if NETSTANDARD
using System.Runtime.Loader;
using Microsoft.Extensions.DependencyModel;
#endif
using System.Text;

namespace Magicodes.ExporterAndImporter.Core.Extension
{
    /// <summary>
    ///     Defines the <see cref="TypeExtensions" />
    /// </summary>
    public static class TypeExtensions
    {
        /// <summary>
        ///     获取显示名
        /// </summary>
        /// <param name="customAttributeProvider"></param>
        /// <param name="inherit"></param>
        /// <returns></returns>
        public static string GetDisplayName(this ICustomAttributeProvider customAttributeProvider, bool inherit = false)
        {
            var displayAttribute = customAttributeProvider.GetAttribute<DisplayAttribute>();
            string displayName;
            if (displayAttribute != null)
            {
                displayName = displayAttribute.Name;
            }
            else
            {
                displayName = customAttributeProvider.GetAttribute<DisplayNameAttribute>()?.DisplayName;
            }
            return displayName;
        }

        /// <summary>
        ///     获取Format
        /// </summary>
        /// <param name="customAttributeProvider"></param>
        /// <returns></returns>
        public static string GetDisplayFormat(this ICustomAttributeProvider customAttributeProvider)
        {
            var formatAttribute = customAttributeProvider.GetAttribute<DisplayFormatAttribute>();
            string displayFormat = string.Empty;
            if (formatAttribute != null)
            {
                displayFormat = formatAttribute.DataFormatString;
            }
            return displayFormat;
        }

        /// <summary>
        ///     获取类型描述
        /// </summary>
        /// <param name="customAttributeProvider"></param>
        /// <param name="inherit"></param>
        /// <returns></returns>
        public static string GetDescription(this ICustomAttributeProvider customAttributeProvider, bool inherit = false)
        {
            var des = string.Empty;
            var desAttribute = customAttributeProvider.GetAttribute<DescriptionAttribute>();
            if (desAttribute != null) des = desAttribute.Description;
            return des;
        }

        /// <summary>
        ///     获取类型描述或显示名
        /// </summary>
        /// <param name="customAttributeProvider"></param>
        /// <param name="inherit"></param>
        /// <returns></returns>
        public static string GetTypeDisplayOrDescription(this ICustomAttributeProvider customAttributeProvider,
            bool inherit = false)
        {
            var displayDescription = customAttributeProvider.GetDescription(inherit);
            if (displayDescription.IsNullOrWhiteSpace()) displayDescription = customAttributeProvider.GetDisplayName(inherit);
            return displayDescription ?? string.Empty;
        }


        /// <summary>
        ///     获取程序集属性
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
        ///     获取程序集属性
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="assembly"></param>
        /// <param name="inherit"></param>
        /// <returns></returns>
        public static IEnumerable<T> GetAttributes<T>(this ICustomAttributeProvider assembly, bool inherit = false)
            where T : Attribute
        {
            return assembly
                .GetCustomAttributes(typeof(T), inherit)
                .OfType<T>();
        }

        /// <summary>
        ///     检查指定指定类型成员中是否存在指定的Attribute特性
        /// </summary>
        /// <typeparam name="T">要检查的Attribute特性类型</typeparam>
        /// <param name="assembly">The assembly<see cref="ICustomAttributeProvider" /></param>
        /// <param name="inherit">是否从继承中查找</param>
        /// <returns>是否存在</returns>
        public static bool AttributeExists<T>(this ICustomAttributeProvider assembly, bool inherit = false)
            where T : Attribute
        {
            return assembly.GetCustomAttributes(typeof(T), inherit).Any(m => m as T != null);
        }

        /// <summary>
        ///     是否必填
        /// </summary>
        /// <param name="propertyInfo"></param>
        /// <returns></returns>
        public static bool IsRequired(this PropertyInfo propertyInfo)
        {
            if (propertyInfo.GetAttribute<RequiredAttribute>(true) != null) return true;
            //Boolean、Byte、SByte、Int16、UInt16、Int32、UInt32、Int64、UInt64、Char、Double、Single
            if (propertyInfo.PropertyType.IsPrimitive) return true;
            switch (propertyInfo.PropertyType.Name)
            {
                case "DateTime":
                case "Decimal":
                    return true;
            }

            return false;
        }

        /// <summary>
        ///     获取当前程序集中应用此特性的类
        /// </summary>
        /// <typeparam name="TAttribute"></typeparam>
        /// <param name="assembly"></param>
        /// <param name="inherit">The inherit<see cref="bool" /></param>
        /// <returns></returns>
        public static IEnumerable<Type> GetTypesWith<TAttribute>(this Assembly assembly, bool inherit)
            where TAttribute : Attribute
        {
            var attrType = typeof(TAttribute);
            foreach (var type in assembly.GetTypes().Where(type => type.GetCustomAttributes(attrType, true).Length > 0))
                yield return type;
        }

        /// <summary>
        ///     获取枚举定义列表
        /// </summary>
        /// <returns>返回枚举列表元组（名称、值、显示名、描述）</returns>
        public static IEnumerable<Tuple<string, int, string, string>> GetEnumDefinitionList(this Type type)
        {
            var list = new List<Tuple<string, int, string, string>>();
            var attrType = type;
            if (!attrType.IsEnum) return null;
            var names = Enum.GetNames(attrType);
            var values = Enum.GetValues(attrType);
            var index = 0;
            foreach (var value in values)
            {
                var name = names[index];
                var field = value.GetType().GetField(value.ToString());
                var displayName = field.GetDisplayName();
                var des = field.GetDescription();
                var item = new Tuple<string, int, string, string>(
                    name,
                    Convert.ToInt32(value),
                    displayName.IsNullOrWhiteSpace() ? null : displayName,
                    des.IsNullOrWhiteSpace() ? null : des
                );
                list.Add(item);
                index++;
            }

            return list;
        }

        /// <summary>
        ///     是否为可为空类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static bool IsNullable(this Type type)
        {
            return Nullable.GetUnderlyingType(type) != null;
        }

        /// <summary>
        ///     获取可为空类型的底层类型
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static Type GetNullableUnderlyingType(this Type type)
        {
            return Nullable.GetUnderlyingType(type);
        }


        /// <summary>
        ///     获取枚举列表
        /// </summary>
        /// <param name="type"></param>
        /// <returns>
        ///     key :返回显示名称或者描述
        ///     value：值
        /// </returns>
        public static IDictionary<string, int> GetEnumTextAndValues(this Type type)
        {
            if (!type.IsEnum) throw new InvalidOperationException();
            var items = type.GetEnumDefinitionList();
            var dic = new Dictionary<string, int>();
            //枚举名 值 显示名称 描述
            foreach (var tuple in items)
            {
                //如果描述、显示名不存在，则返回枚举名称
                dic.Add(tuple.Item4 ?? tuple.Item3 ?? tuple.Item1, tuple.Item2);
            }
            return dic;
        }

        /// <summary>
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static string GetCSharpTypeName(this Type type)
        {
            var sb = new StringBuilder();
            var name = type.Name;
            if (!type.IsGenericType) return name;
            sb.Append(name.Substring(0, name.IndexOf('`')));
            sb.Append("<");
            sb.Append(string.Join(", ", type.GetGenericArguments()
                .Select(t => t.GetCSharpTypeName())));

            sb.Append(">");
            return sb.ToString();
        }

        /// <summary>
        ///     实例化依赖
        /// </summary>
        /// <param name="type"></param>
        /// <returns></returns>
        public static object[] CreateType(this Type type)
        {
            //Get the first
            var constructorInfo = type.GetConstructors().FirstOrDefault();
            var parameterInfos = constructorInfo?.GetParameters();
            var objects = new List<object>();
            //GetAssemblies need to add conditional screening
            //var getAssemblies = AppDomain.CurrentDomain.GetAssemblies();
            List<Type> types = new List<Type>();
            foreach (var item in GetAllAssemblies())
            {
                try
                {
                    foreach (var typeItem in item.GetTypes())
                    {
                        types.Add(typeItem);
                    }
                }
                catch (Exception)
                {
                    // ignored
                }

            }
            foreach (var item in parameterInfos)
            {
                var t = types.FirstOrDefault(x => x.GetInterfaces().Any(a => a.Name == item.ParameterType.Name));
                var obj = Activator.CreateInstance(t, CreateType(t));
                objects.Add(obj);
            }
            return objects.ToArray();
        }

        /// <summary>
        /// 获取项目程序集,排除所有的系统程序集(Microsoft.***、System.***等)、Nuget包
        /// </summary>
        /// <returns></returns>
        public static IList<Assembly> GetAllAssemblies()
        {
#if NETSTANDARD
            var list = new List<Assembly>();
            var deps = DependencyContext.Default;
            //var libs = deps.CompileLibraries.Where(lib => !lib.Serviceable && lib.Type != "package");
            foreach (var lib in deps.CompileLibraries)
            {
                try
                {
                    var assembly = AssemblyLoadContext.Default.LoadFromAssemblyName(new AssemblyName(lib.Name));
                    list.Add(assembly);
                }
                catch (Exception)
                {
                    // ignored
                }
            }
            return list;
#else
            return AppDomain.CurrentDomain.GetAssemblies();
#endif

        }

        /// <summary>
        ///     获取私有属性值
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="instance"></param>
        /// <param name="propertyname"></param>
        /// <returns></returns>
        public static T GetPrivateProperty<T>(this object instance, string propertyname)
        {
            Type type = instance.GetType().BaseType;
            FieldInfo[] finfos = type.GetFields(BindingFlags.Instance | BindingFlags.NonPublic);
            var field = finfos.FirstOrDefault(f => f.Name == propertyname);
            return (T)field.GetValue(instance);
        }

        /// <summary>
        /// 获取筛选器
        /// </summary>
        /// <typeparam name="TFilter"></typeparam>
        /// <param name="filterType"></param>
        /// <param name="isDisableAllFilter"></param>
        /// <returns></returns>
        public static TFilter GetFilter<TFilter>(this Type filterType, bool isDisableAllFilter) where TFilter : IFilter
        {
            TFilter filter = default;

            if (!isDisableAllFilter)
            {
#if NETSTANDARD
                //判断容器中是否已注册
                if (AppDependencyResolver.HasInit)
                {
                    filter = AppDependencyResolver.Current.GetService<TFilter>();
                }
                else if (filterType != null && typeof(TFilter).IsAssignableFrom(filterType))
                {
                    filter = (TFilter)filterType.Assembly.CreateInstance(filterType.FullName, true, System.Reflection.BindingFlags.Default, null, filterType.CreateType(), null, null);
                }

#else
                if (filterType != null && typeof(TFilter).IsAssignableFrom(filterType))
                {
                    filter = (TFilter)filterType.Assembly.CreateInstance(filterType.FullName, true,
                        System.Reflection.BindingFlags.Default, null, filterType.CreateType(), null, null);
                }
#endif
            }
            return filter;
        }
    }
}