// ======================================================================
// 
//           filename : IExporterByTemplate.cs
//           description :
// 
//           created by 雪雁 at  2019-09-26 14:59
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

namespace Magicodes.ExporterAndImporter.Core
{
    /// <summary>
    /// 根据模板导出
    /// </summary>
    public interface IExporterByTemplate : IExportListFileByTemplate, IExportListStringByTemplate, IExportStringByTemplate, IExportFileByTemplate
    {

    }
}