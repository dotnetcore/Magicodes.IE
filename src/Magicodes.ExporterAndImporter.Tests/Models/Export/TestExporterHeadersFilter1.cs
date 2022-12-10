// ======================================================================
// 
//           filename : AttrsLocalizationTestData.cs
//           description :
// 
//           created by 雪雁 at  2019-11-05 20:02
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using Magicodes.ExporterAndImporter.Core.Filters;
using Magicodes.ExporterAndImporter.Core.Models;
using System.Collections.Generic;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    public class TestExporterHeadersFilter1 : IExporterHeadersFilter
    {
        public IList<ExporterHeaderInfo> Filter(IList<ExporterHeaderInfo> exporterHeaderInfos)
        {
            IList<ExporterHeaderInfo> filtered = new List<ExporterHeaderInfo>();
            foreach (var item in exporterHeaderInfos)
            {
                if (item.PropertyName == "Text2")
                {
                    item.DisplayName = "标题";
                    item.ExporterHeaderAttribute.ColumnIndex = -1;//排到第一列
                }
            }
            return exporterHeaderInfos;
        }
    }
}