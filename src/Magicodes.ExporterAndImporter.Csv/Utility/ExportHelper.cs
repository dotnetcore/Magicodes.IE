using CsvHelper;
using Magicodes.ExporterAndImporter.Core.Extension;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Magicodes.ExporterAndImporter.Core;

namespace Magicodes.ExporterAndImporter.Csv.Utility
{
    /// <summary>
    ///     Csv导出辅助类
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ExportHelper<T> where T : class
    {
        private readonly Type _type;

        /// <summary>
        /// </summary>
        public ExportHelper()
        {
        }

        /// <summary>
        /// </summary>
        /// <param name="type"></param>
        public ExportHelper(Type type)
        {
            this._type = type;
        }

        /// <summary>
        ///     导出Csv
        /// </summary>
        /// <param name="dataItems"></param>
        /// <returns></returns>
        public byte[] GetCsvExportAsByteArray(ICollection<T> dataItems = null, string delimiter = "")
        {
            using (var ms = new MemoryStream())
            using (var writer = new StreamWriter(ms, Encoding.UTF8))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.Context.Configuration.HasHeaderRecord = true;

                if(!string.IsNullOrWhiteSpace(delimiter))
                    csv.Context.Configuration.Delimiter = delimiter;

                if (_type == null)
                {
                    csv.Context.RegisterClassMap<AutoMap<T>>();
                }
                else
                {
                    csv.Context.RegisterClassMap<AutoMap<T>>();
                }

                if (dataItems != null && dataItems.Count > 0)
                {
                    csv.WriteRecords(dataItems);
                }
                writer.Flush();
                ms.Position = 0;
                return ms.ToArray();
            }
        }

        /// <summary>
        ///     导出表头
        /// </summary>
        /// <returns></returns>
        public byte[] GetCsvExportHeaderAsByteArray(string delimiter = "")
        {
            var properties = typeof(T).GetSortedPropertyInfos();

            using (var ms = new MemoryStream())
            using (var writer = new StreamWriter(ms, Encoding.UTF8))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.Context.Configuration.HasHeaderRecord = true;

                if (!string.IsNullOrWhiteSpace(delimiter))
                    csv.Context.Configuration.Delimiter = delimiter;

                #region header

                foreach (var prop in properties)
                {
                    var name = prop.Name;
                    var headerAttribute = prop.GetCustomAttribute<Core.ExporterHeaderAttribute>();
                    if (headerAttribute != null)
                    {
                        name = headerAttribute.DisplayName ?? prop.GetDisplayName() ?? prop.Name;
                    }
                    var importAttribute = prop.GetCustomAttribute<Core.ImporterHeaderAttribute>();
                    if (importAttribute != null)
                    {
                        name = importAttribute.Name ?? prop.GetDisplayName() ?? prop.Name;
                    }
                    csv.WriteField(name);
                }
                csv.NextRecord();
                #endregion

                writer.Flush();
                ms.Position = 0;
                return ms.ToArray();
            }
        }

        /// <summary>
        ///     导出Csv
        /// </summary>
        /// <param name="dataItems"></param>
        /// <returns></returns>
        public byte[] GetCsvExportAsByteArray(DataTable dataItems, string delimiter = "")
        {
            using (var ms = new MemoryStream())
            using (var writer = new StreamWriter(ms, Encoding.UTF8))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.Context.RegisterClassMap<AutoMap<T>>();
                csv.Context.Configuration.HasHeaderRecord = true;

                if (!string.IsNullOrWhiteSpace(delimiter))
                    csv.Context.Configuration.Delimiter = delimiter;

                //#region header 
                //var properties = typeof(T).GetProperties();
                //foreach (var prop in properties)
                //{
                //    var name = prop.Name;
                //    var headerAttribute = prop.GetCustomAttribute<Core.ExporterHeaderAttribute>();
                //    if (headerAttribute != null)
                //    {
                //        name = headerAttribute.DisplayName ?? prop.GetDisplayName() ?? prop.Name;
                //    }
                //    var importAttribute = prop.GetCustomAttribute<Core.ImporterHeaderAttribute>();
                //    if (importAttribute != null)
                //    {
                //        name = importAttribute.Name ?? prop.GetDisplayName() ?? prop.Name;
                //    }
                //    csv.WriteField(name);
                //}
                //csv.NextRecord();
                //#endregion
                //foreach (DataRow row in dataItems.Rows)
                //{
                //    for (var i = 0; i < dataItems.Columns.Count; i++)
                //    {
                //        csv.WriteField(row[i]);
                //    }
                //    csv.NextRecord();
                //}
                csv.WriteRecords(dataItems.ToList<T>());
                writer.Flush();
                ms.Position = 0;
                return ms.ToArray();
            }
        }
    }
}
