using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class TestExporterHeaderFilter1 : IExporterHeaderFilter
    {
        /// <summary>
        /// 表头筛选器（修改名称）
        /// </summary>
        /// <param name="exporterHeaderInfo"></param>
        /// <returns></returns>
        public ExporterHeaderInfo Filter(ExporterHeaderInfo exporterHeaderInfo)
        {
            if (exporterHeaderInfo.DisplayName.Equals("名称"))
            {
                exporterHeaderInfo.DisplayName = "name";
            }
            return exporterHeaderInfo;
        }
    }

    public class TestExporterHeaderFilter2 : IExporterHeaderFilter
    {
        /// <summary>
        /// 表头筛选器（修改忽略列）
        /// </summary>
        /// <param name="exporterHeaderInfo"></param>
        /// <returns></returns>
        public ExporterHeaderInfo Filter(ExporterHeaderInfo exporterHeaderInfo)
        {
            if (exporterHeaderInfo.ExporterHeaderAttribute.IsIgnore)
            {
                exporterHeaderInfo.ExporterHeaderAttribute.IsIgnore = false;
            }
            return exporterHeaderInfo;
        }
    }

    public class DataTableTestExporterHeaderFilter : IExporterHeaderFilter
    {
        /// <summary>
        /// 表头筛选器（修改忽略列）
        /// </summary>
        /// <param name="exporterHeaderInfo"></param>
        /// <returns></returns>
        public ExporterHeaderInfo Filter(ExporterHeaderInfo exporterHeaderInfo)
        {
            if (exporterHeaderInfo.DisplayName.Equals("Number"))
            {
                exporterHeaderInfo.DisplayName = "数值";
            }
            return exporterHeaderInfo;
        }
    }
}
