// ======================================================================
// 
//           filename : DataTableHelper.cs
//           description :
// 
//           created by 雪雁 at  2019-11-23 19:45
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using System.Data;

namespace Magicodes.ExporterAndImporter.Excel.Utility
{
    public static class DataTableHelper
    {
        /// <summary>
        ///     分解数据表
        /// </summary>
        /// <param name="originalTab">需要分解的表</param>
        /// <param name="rowsNum">每个表包含的数据量</param>
        /// <returns></returns>
        public static DataSet SplitDataTable(this DataTable originalTab, int rowsNum = 1000000)
        {
            //获取所需创建的表数量
            var tableNum = originalTab.Rows.Count / rowsNum;

            //获取数据余数
            var remainder = originalTab.Rows.Count % rowsNum;

            if (remainder != 0) tableNum += 1;

            var ds = new DataSet();

            //如果只需要创建1个表，直接将原始表存入DataSet
            if (tableNum == 1)
            {
                ds.Tables.Add(originalTab);
            }
            else
            {
                var tableSlice = new DataTable[tableNum];

                //Save orginal columns into new table.            
                for (var c = 0; c < tableNum; c++)
                {
                    tableSlice[c] = new DataTable();
                    foreach (DataColumn dc in originalTab.Columns)
                        tableSlice[c].Columns.Add(dc.ColumnName, dc.DataType);
                }

                //Import Rows
                for (var i = 0; i < tableNum; i++)
                    if (remainder == 0)
                    {
                        for (var j = i * rowsNum; j < (i + 1) * rowsNum; j++)
                            tableSlice[i].ImportRow(originalTab.Rows[j]);
                    }
                    else
                    {
                        // if the current table is not the last one
                        if (i != tableNum - 1)
                            for (var j = i * rowsNum; j < (i + 1) * rowsNum; j++)
                                tableSlice[i].ImportRow(originalTab.Rows[j]);
                        else
                            for (var k = i * rowsNum; k < i * rowsNum + remainder; k++)
                                tableSlice[i].ImportRow(originalTab.Rows[k]);
                    }

                //add all tables into a dataset                
                foreach (var dt in tableSlice) ds.Tables.Add(dt);
            }

            return ds;
        }
    }
}