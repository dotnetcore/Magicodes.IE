using Magicodes.ExporterAndImporter.Core.Extension;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Magicodes.ExporterAndImporter.Core.Models
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="TData"></typeparam>
    public class ExportDocumentInfo<TData> where TData : class
    {
        /// <summary>
        /// 
        /// </summary>
        public ExportDocumentInfo(IList<TData> datas)
        {
            Headers = new List<ExporterHeaderAttribute>();
            Datas = datas;
            Title = typeof(TData).GetAttribute<ExporterAttribute>()?.Name ?? typeof(TData).Name;

            foreach (var propertyInfo in typeof(TData).GetProperties())
            {
                var exporterHeader = propertyInfo.PropertyType.GetAttribute<ExporterHeaderAttribute>() ?? new ExporterHeaderAttribute()
                {
                    DisplayName = propertyInfo.GetDisplayName() ?? propertyInfo.Name
                };
                Headers.Add(exporterHeader);
            }
        }


        /// <summary>
        /// 文档标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// 头部信息
        /// </summary>
        public IList<ExporterHeaderAttribute> Headers { get; set; }

        /// <summary>
        /// 数据
        /// </summary>
        public IList<TData> Datas { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public DataTable ToDataTable()
        {
            var props = typeof(TData).GetProperties();
            var dt = new DataTable();
            dt.Columns.AddRange(props.Select(p => new DataColumn(p.PropertyType.GetAttribute<ExporterAttribute>()?.Name ?? p.GetDisplayName() ?? p.Name, p.PropertyType)).ToArray());
            if (Datas.Count <= 0)
            {
                return dt;
            }

            for (var i = 0; i < Datas.Count; i++)
            {
                var tempList = new ArrayList();
                foreach (var obj in props.Select(pi => pi.GetValue(Datas.ElementAt(i), null)))
                {
                    tempList.Add(obj);
                }
                var array = tempList.ToArray();
                dt.LoadDataRow(array, true);
            }
            return dt;
        }
    }
}
