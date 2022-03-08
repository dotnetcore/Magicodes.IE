using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.NetworkInformation;
using System.Reflection;

namespace Magicodes.ExporterAndImporter.Core.Extension
{
    public static class ImporterHeaderAttributeExtensions
    {
        public static PropertyInfo[] GetSortedPropertyInfos(this Type t)
        {
            var props = t.GetProperties();
            var noIndex = new List<PropertyInfo>();
            var hasIndex = new Dictionary<int, PropertyInfo>();
            var result = new PropertyInfo[props.Length];
            foreach (var propertyInfo in props)
            {
                var index = propertyInfo.GetAttribute<ImporterHeaderAttribute>()?.ColumnIndex;
                if (index != null && index != 0)
                {
                    hasIndex.Add(index.Value, propertyInfo);
                }
                else
                {
                    noIndex.Add(propertyInfo);
                }
            }

            for (var i = 0; i < props.Length; i++)
            {
                if (hasIndex.ContainsKey(i + 1))
                {
                    result[i] = hasIndex[i + 1];
                    hasIndex.Remove(i + 1);
                }
                else
                {
                    var firstNoIndex = noIndex.FirstOrDefault();
                    if (firstNoIndex == null)
                    {
                        var minIndex = hasIndex.Keys.Min();
                        result[i] = hasIndex[minIndex];
                        hasIndex.Remove(minIndex);
                    }
                    else
                    {
                        result[i] = firstNoIndex;
                        noIndex.Remove(firstNoIndex);
                    }
                }
            }

            return result;
        }
    }
}