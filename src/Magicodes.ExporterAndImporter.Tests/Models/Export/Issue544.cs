using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using Magicodes.IE.Core;
using System;
using System.Collections.Generic;
using System.Reflection;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    [ExcelExporter(Name = "导出结果", TableStyle = OfficeOpenXml.Table.TableStyles.None)]
    public class Issue544
    {
        /// <summary>
        /// 名称
        /// </summary>
        [ExporterHeader(DisplayName = "姓名")]
        public string Name { get; set; }

        /// <summary>
        /// 性别
        /// </summary>
        [ExporterHeader(DisplayName = "性别")]
        public string Gender { get; set; }

        /// <summary>
        /// 是否校友
        /// </summary>
        [ExporterHeader(DisplayName = "是否校友")]
        [BoolLocal337]
        public bool? IsAlumni { get; set; }

        [ExporterHeader(DisplayName = "是否校友2")]
        [BoolLocal337]
        public bool IsAlumni2 { get; set; }
    }


    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
    public class BoolLocal337Attribute : ValueMappingsBaseAttribute
    {
        public override Dictionary<string, object> GetMappings(PropertyInfo propertyInfo)
        {
            var res= new Dictionary<string, object>();
            res.Add("是",true);
            res.Add("否",false);
            return res;
        }
    }
}
