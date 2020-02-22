using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.TypeConversion;
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Extension;
using System;
using System.Linq;
using System.Reflection;

namespace Magicodes.ExporterAndImporter.Csv
{
    /// <summary>
    ///     EnumConverter
    /// </summary>
    public class CsvHelperEnumConverter : DefaultTypeConverter
    {
        /// <summary>
        ///     从字符串反转
        /// </summary>  
        /// <param name="text"></param>
        /// <param name="row"></param>
        /// <param name="memberMapData"></param>
        /// <returns></returns>
        public override object ConvertFromString(string text, IReaderRow row, MemberMapData memberMapData)
        {
            var type = memberMapData.Member.MemberType();
            var value = memberMapData.Member.GetCustomAttributes<ValueMappingAttribute>().FirstOrDefault(f => f.Text == text)?.Value;
            var isNullable = type.IsNullable();
            if (isNullable) type = type.GetNullableUnderlyingType();
            var values = type.GetEnumTextAndValues();

            if (value == null)
            {
                value = Enum.ToObject(type, values.FirstOrDefault(f => f.Key == text).Value);
            }
            return value ?? text;
        }

    }
}
