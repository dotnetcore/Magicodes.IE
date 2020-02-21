using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using CsvHelper;

namespace Magicodes.ExporterAndImporter.Csv.Utility
{
    /// <summary>
    ///     Csv导出辅助类
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ExportHelper<T> where T:class
    {
        /// <summary>
        ///     导出CSV
        /// </summary>
        /// <param name="dataItems"></param>
        /// <returns></returns>
        public byte[] GetCsvExportAsByteArray(ICollection<T> dataItems=null)
        {
            using (var ms = new MemoryStream())
            using (var writer = new StreamWriter(ms, Encoding.UTF8))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.Configuration.HasHeaderRecord = true;
                csv.Configuration.RegisterClassMap<AutoMap<T>>();
                csv.WriteRecords(dataItems);
                writer.Flush();
                ms.Position = 0;
                return ms.ToArray();
            }
        }

    }
}
