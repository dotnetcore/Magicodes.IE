using CsvHelper;
using CsvHelper.Configuration;
using CsvHelper.TypeConversion;
using Magicodes.ExporterAndImporter.Core;
using System.Linq;
using System.Reflection;

namespace Magicodes.ExporterAndImporter.Excel
{
    /// <summary>
    ///     EnumConverter
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class CsvHelperEnumConverter<T> : DefaultTypeConverter where T : struct
    {
        /// <summary>
        ///     
        /// </summary>
        /// <param name="value"></param>
        /// <param name="row"></param>
        /// <param name="memberMapData"></param>
        /// <returns></returns>
        public override string ConvertToString(object value, IWriterRow row, MemberMapData memberMapData)
        {
            return base.ConvertToString(value, row, memberMapData);
        }
        /// <summary>
        ///     从字符串反转
        /// </summary>
        /// <param name="text"></param>
        /// <param name="row"></param>
        /// <param name="memberMapData"></param>
        /// <returns></returns>
        public override object ConvertFromString(string text, IReaderRow row, MemberMapData memberMapData)
        {
            var type = memberMapData.Member;

            var mappings = type.GetCustomAttributes<ValueMappingAttribute>().ToList();

            return mappings.FirstOrDefault(f => f.Text == text)?.Value;

        }
    }
}
