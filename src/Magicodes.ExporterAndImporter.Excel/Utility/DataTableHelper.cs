using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace Magicodes.ExporterAndImporter.Excel.Utility
{
    public static class  DataTableHelper
    {

        /// <summary>
        /// 分解数据表
        /// </summary>
        /// <param name="originalTab">需要分解的表</param>
        /// <param name="rowsNum">每个表包含的数据量</param>
        /// <returns></returns>
        public static DataSet SplitDataTable(this DataTable originalTab, int rowsNum = 1000000)
        {
            //获取所需创建的表数量
            int tableNum = originalTab.Rows.Count / rowsNum;

            //获取数据余数
            int remainder = originalTab.Rows.Count % rowsNum;

            if(remainder !=0)
            {
                tableNum += 1;
            }

            DataSet ds = new DataSet();

            //如果只需要创建1个表，直接将原始表存入DataSet
            if (tableNum == 0)
            {
                ds.Tables.Add(originalTab);
            }
            else
            {
                DataTable[] tableSlice = new DataTable[tableNum];

                //Save orginal columns into new table.            
                for (int c = 0; c < tableNum; c++)
                {
                    tableSlice[c] = new DataTable();
                    foreach (DataColumn dc in originalTab.Columns)
                    {
                        tableSlice[c].Columns.Add(dc.ColumnName, dc.DataType);
                    }
                }
                //Import Rows
                for (int i = 0; i < tableNum; i++)
                {
                    if (remainder == 0)
                    {
                        for (int j = i * rowsNum; j < ((i + 1) * rowsNum); j++)
                        {
                            tableSlice[i].ImportRow(originalTab.Rows[j]);
                        }
                    }
                    else
                    { 
                        // if the current table is not the last one
                        if (i != tableNum - 1)
                        {
                            for (int j = i * rowsNum; j < ((i + 1) * rowsNum); j++)
                            {
                                tableSlice[i].ImportRow(originalTab.Rows[j]);
                            }
                        }
                        else
                        {
                            for (int k = i * rowsNum; k < (i * rowsNum + remainder); k++)
                            {
                                tableSlice[i].ImportRow(originalTab.Rows[k]);
                            }
                        }
                    }
                }

                //add all tables into a dataset                
                foreach (DataTable dt in tableSlice)
                {
                    ds.Tables.Add(dt);
                }
            }
            return ds;
        }
    }
}
