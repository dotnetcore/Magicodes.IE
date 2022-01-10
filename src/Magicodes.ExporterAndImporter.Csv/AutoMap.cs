using CsvHelper.Configuration;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using System.Globalization;
using System.Reflection;

namespace Magicodes.ExporterAndImporter.Csv
{
    /// <summary>
    ///     动态构建映射
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class AutoMap<T> : ClassMap<T>
    {
        /// <summary>
        ///     构造方法
        /// </summary>
        public AutoMap()
        {
            var properties = typeof(T).GetProperties();
            foreach (var prop in properties)
            {
                var result = MapProperty(prop);
                var tcOption = result.Item1.TypeConverterOption;
                tcOption.NumberStyles(NumberStyles.Any);
                tcOption.DateTimeStyles(DateTimeStyles.None);
                var format = tcOption.Format();
                if (result.Item2 != null)
                {
                    if (!string.IsNullOrEmpty(result.Item2.Format))
                    {
                        tcOption.Format(result.Item2.Format);
                    }
                    if (result.Item2.IsIgnore)
                    {
                        format.Ignore();
                    }
                }
                else if (result.Item3 != null)
                {
                    if (result.Item3.IsIgnore)
                    {
                        format.Ignore();
                    }
                }

                if (prop.PropertyType.IsEnum)
                {
                    result.Item1.TypeConverter<CsvHelperEnumConverter>();
                }
                else
                {
                    var isNullable = prop.PropertyType.IsNullable();
                    if (!isNullable) continue;
                    var type = prop.PropertyType.GetNullableUnderlyingType();
                    if (type.IsEnum)
                    {
                        result.Item1.TypeConverter<CsvHelperEnumConverter>();
                    }
                }
            }
        }
        /// <summary>
        /// </summary>
        /// <param name="property"></param>
        /// <returns></returns>
        private (MemberMap, ExporterHeaderAttribute, ImporterHeaderAttribute) MapProperty(PropertyInfo property)
        {
            var map = Map(typeof(T), property);
            var name = property.Name;
            var headerAttribute = property.GetCustomAttribute<ExporterHeaderAttribute>();
            if (headerAttribute != null)
            {
                name = headerAttribute.DisplayName ?? property.GetDisplayName() ?? property.Name;
            }
            var importAttribute = property.GetCustomAttribute<ImporterHeaderAttribute>();
            if (importAttribute != null)
            {
                name = importAttribute.Name ?? property.GetDisplayName() ?? property.Name;
            }
            map.Name(name);
            return (map, headerAttribute, importAttribute);
        }
    }
}