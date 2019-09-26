// ======================================================================
// 
//           Copyright (C) 2019-2030 湖南心莱信息科技有限公司
//           All rights reserved
// 
//           filename : Extension.cs
//           description :
// 
//           created by 雪雁 at  2019-09-26 13:51
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Magicodes.ExporterAndImporter.Core.Extension
{
    /// <summary>
    /// 
    /// </summary>
    public static class ListExtension
    {
        /// <summary>
        /// 将List集合转成DataTable
        /// </summary>
        /// <returns></returns>
        public static DataTable ToDataTable<T>(this IList<T> list)
        {
            var props = typeof(T).GetProperties();
            var dt = new DataTable();
            dt.Columns.AddRange(props.Select(p =>
                new DataColumn(p.PropertyType.GetAttribute<ExporterAttribute>()?.Name ?? p.GetDisplayName() ?? p.Name,
                    p.PropertyType)).ToArray());
            if (list.Count <= 0) return dt;

            for (var i = 0; i < list.Count; i++)
            {
                var tempList = new ArrayList();
                foreach (var obj in props.Select(pi => pi.GetValue(list.ElementAt(i), null))) tempList.Add(obj);
                var array = tempList.ToArray();
                dt.LoadDataRow(array, true);
            }

            return dt;
        }
    }
}