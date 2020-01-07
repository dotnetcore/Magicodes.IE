// ======================================================================
// 
//           filename : TemplateExportResult.cs
//           description :
// 
//           created by 雪雁 at  -- 
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using System;
using System.Collections.Generic;

namespace Magicodes.ExporterAndImporter.Excel.Utility.TemplateExport
{
    /// <summary>
    ///     模板导出结果
    /// </summary>
    public class TemplateExportResult
    {
        /// <summary>
        /// </summary>
        public TemplateExportResult()
        {

        }

        /// <summary>
        ///     模板错误
        /// </summary>
        public virtual IList<TemplateFieldErrorInfo> TemplateErrors { get; set; }

        /// <summary>
        ///     异常信息
        /// </summary>
        public virtual Exception Exception { get; set; }

        /// <summary>
        ///     是否存在导入错误
        /// </summary>
        public virtual bool HasError => Exception != null || (TemplateErrors?.Count > 0);
    }
}