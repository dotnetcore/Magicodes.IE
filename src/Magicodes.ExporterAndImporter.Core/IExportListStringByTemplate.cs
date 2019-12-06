// ======================================================================
// 
//           filename : IExportListStringByTemplate.cs
//           description :
// 
//           created by 雪雁 at  -- 
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================


using System.Collections.Generic;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Core
{
    /// <summary>
    /// 根据模板导出列表字符串
    /// </summary>
    public interface IExportListStringByTemplate
    {
        /// <summary>
        ///     根据模板导出列表
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="dataItems"></param>
        /// <param name="htmlTemplate">Html模板内容</param>
        /// <returns></returns>
        Task<string> ExportListByTemplate<T>(ICollection<T> dataItems, string htmlTemplate = null) where T : class;
    }
}