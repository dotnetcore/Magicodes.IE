using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace Magicodes.ExporterAndImporter.Excel.Utility
{
    public static class EnumHelper
    {
        public static IDictionary<string, int> GetDisplayNames(Type type)
        {
            if (!type.IsEnum) throw new InvalidOperationException();
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


        public static T? GetNullableValue<T>(string displayName) where T : struct
        {
            if (string.IsNullOrEmpty(displayName))
            {
                return null;
            }
            return GetValue<T>(displayName);
        }

        public static T GetValue<T>(string displayName) where T : struct
        {
            var type = typeof(T);
            if (!type.IsEnum) throw new InvalidOperationException();
            foreach (var field in type.GetFields())
            {
                var displayAttribute = field.GetCustomAttributes(typeof(DisplayAttribute), false)
                    .SingleOrDefault() as DisplayAttribute;
                if (displayAttribute != null && displayAttribute.Name == displayName)
                {
                    return (T)field.GetValue(null);
                }
            }
            return default;
        }

        public static bool IsValid<T>(string displayName)
        {
            T value = default;
            return TryGetValue(displayName, ref value);
        }

        public static bool TryGetValue<T>(string displayName, ref T value)
        {
            var type = typeof(T);
            if (!type.IsEnum) throw new InvalidOperationException();
            foreach (var field in type.GetFields())
            {
                var displayAttribute = field.GetCustomAttributes(typeof(DisplayAttribute), false)
                    .SingleOrDefault() as DisplayAttribute;
                if (displayAttribute != null && displayAttribute.Name == displayName)
                {
                    value = (T)field.GetValue(null);
                    return true;
                }
            }
            return false;
        }
    }
}