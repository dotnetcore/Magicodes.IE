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
    }
}