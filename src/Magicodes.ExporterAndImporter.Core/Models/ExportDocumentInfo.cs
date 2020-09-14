// ======================================================================
// 
//           filename : ExportDocumentInfo.cs
//           description :
// 
//           created by 雪雁 at  2019-09-26 14:59
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using Magicodes.ExporterAndImporter.Core.Extension;
using System;
using System.Collections.Generic;
using System.Data;

namespace Magicodes.ExporterAndImporter.Core.Models
{
    /// <summary>
    /// </summary>
    /// <typeparam name="TData"></typeparam>
    public class ExportDocumentInfoOfListData<TData> where TData : class
    {
        /// <summary>
        /// </summary>
        public ExportDocumentInfoOfListData(ICollection<TData> datas)
        {
            Headers = new List<ExporterHeaderAttribute>();
            Datas = datas;
            Title = typeof(TData).GetAttribute<ExporterAttribute>()?.Name ?? typeof(TData).Name;

            foreach (var propertyInfo in typeof(TData).GetProperties())
            {
                var exporterHeader = propertyInfo.GetAttribute<ExporterHeaderAttribute>() ??
                                     new ExporterHeaderAttribute
                                     {
                                         DisplayName = propertyInfo.GetDisplayName() ?? propertyInfo.Name
                                     };
                Headers.Add(exporterHeader);
            }
        }


        /// <summary>
        ///     文档标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        ///     头部信息
        /// </summary>
        public IList<ExporterHeaderAttribute> Headers { get; set; }

        /// <summary>
        ///     数据
        /// </summary>
        public ICollection<TData> Datas { get; set; }

        /// <summary>
        /// </summary>
        /// <returns></returns>
        public DataTable ToDataTable()
        {
            return Datas.ToDataTable();
        }
    }
    /// <summary>
    /// </summary>
    public class ExportDocumentInfo
    {
        /// <summary>
        /// </summary>
        public ExportDocumentInfo(object data, Type type)
        {
            Headers = new List<ExporterHeaderAttribute>();
            Data = data;
            Title = type.GetAttribute<ExporterAttribute>()?.Name ?? type.Name;

            foreach (var propertyInfo in type.GetProperties())
            {
                var exporterHeader = propertyInfo.PropertyType.GetAttribute<ExporterHeaderAttribute>() ??
                                     new ExporterHeaderAttribute
                                     {
                                         DisplayName = propertyInfo.GetDisplayName() ?? propertyInfo.Name
                                     };
                Headers.Add(exporterHeader);
            }
        }


        /// <summary>
        ///     文档标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        ///     头部信息
        /// </summary>
        public IList<ExporterHeaderAttribute> Headers { get; set; }

        /// <summary>
        ///     数据
        /// </summary>
        public object Data { get; set; }
    }

    /// <summary>
    /// </summary>
    /// <typeparam name="TData"></typeparam>
    public class ExportDocumentInfo<TData> where TData : class
    {
        /// <summary>
        /// </summary>
        public ExportDocumentInfo(TData data)
        {
            Headers = new List<ExporterHeaderAttribute>();
            Data = data;
            Title = typeof(TData).GetAttribute<ExporterAttribute>()?.Name ?? typeof(TData).Name;

            foreach (var propertyInfo in typeof(TData).GetProperties())
            {
                var exporterHeader = propertyInfo.PropertyType.GetAttribute<ExporterHeaderAttribute>() ??
                                     new ExporterHeaderAttribute
                                     {
                                         DisplayName = propertyInfo.GetDisplayName() ?? propertyInfo.Name
                                     };
                Headers.Add(exporterHeader);
            }
        }


        /// <summary>
        ///     文档标题
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        ///     头部信息
        /// </summary>
        public IList<ExporterHeaderAttribute> Headers { get; set; }

        /// <summary>
        ///     数据
        /// </summary>
        public TData Data { get; set; }
    }
}