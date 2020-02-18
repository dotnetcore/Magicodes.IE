using CsvHelper.Configuration;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace Magicodes.ExporterAndImporter.Excel
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
            var nameProperty = properties.FirstOrDefault(p => p.Name == "Name");
            //if (nameProperty != null)
            //    MapProperty(nameProperty).Item1.Index(0);
            foreach (var prop in properties.Where(p => p != nameProperty))
            {
                var result = MapProperty(prop);
                var tcOption = result.Item1.TypeConverterOption;
                var format = tcOption.Format();
                if (!string.IsNullOrEmpty(result.Item2?.Format))
                {
                    tcOption.Format(result.Item2.Format);
                }
                tcOption.NumberStyles(NumberStyles.Any);
                tcOption.DateTimeStyles(DateTimeStyles.None);
                if (result.Item2?.IsIgnore != null && result.Item2.IsIgnore == true)
                {
                    format.Ignore();
                }

            }
        }

        private (MemberMap, ExporterHeaderAttribute) MapProperty(PropertyInfo property)
        {
            var map = Map(typeof(T), property);
            string name = property.Name;
            var headerAttribute = property.GetCustomAttribute<ExporterHeaderAttribute>();
            if (headerAttribute != null)
            {
                name = headerAttribute.DisplayName ?? property.GetDisplayName() ?? property.Name;
            }
            map.Name(name);
            return (map, headerAttribute);

        }
    }
}