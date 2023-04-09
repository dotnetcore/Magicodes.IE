using OfficeOpenXml;
using System;
using System.Reflection;

namespace Magicodes.IE.Stash.Extensions
{
    /// <summary>
    /// EPPlus扩展
    /// <para>是否应该合并到上层项目中??</para>
    /// </summary>
    public static class EPPlusExtensions
    {
        public static object GetCellValueByType(this ExcelRange excelRange, Type toType)
        {
            var value = excelRange.Value;
            return value.GetCellValueByType(toType);
        }
        public static object GetCellValueByType(this object value, Type toType)
        {
            if (value == null)
            {
                return toType.IsValueType ? Activator.CreateInstance(toType) : null; ;
            }

            var fromType = value.GetType();

            var toNullableUnderlyingType = toType.IsGenericType && toType.GetGenericTypeDefinition() == typeof(Nullable<>)
                ? Nullable.GetUnderlyingType(toType)
                : null;

            if (fromType == toType || fromType == toNullableUnderlyingType)
                return value;

            // if converting to nullable struct and input is blank string, return null
            if (toNullableUnderlyingType != null && fromType == typeof(string) && ((string)value).Trim() == string.Empty)
                return Activator.CreateInstance(toType);

            toType = toNullableUnderlyingType ?? toType;

            if (toType == typeof(DateTime))
            {
                if (value is double)
                    return DateTime.FromOADate((double)value);

                if (fromType == typeof(TimeSpan))
                    return new DateTime(((TimeSpan)value).Ticks);

                if (fromType == typeof(string))
                    return DateTime.Parse(value.ToString());
            }
            else if (toType == typeof(TimeSpan))
            {
                if (value is double)
                    return new TimeSpan(DateTime.FromOADate((double)value).Ticks);

                if (fromType == typeof(DateTime))
                    return new TimeSpan(((DateTime)value).Ticks);

                if (fromType == typeof(string))
                    return TimeSpan.Parse(value.ToString());
            }

            return Convert.ChangeType(value, toType);

        }

    }
}
