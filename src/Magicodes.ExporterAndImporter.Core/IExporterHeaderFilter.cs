using Magicodes.ExporterAndImporter.Core.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace Magicodes.ExporterAndImporter.Core
{
    /// <summary>
    /// 列头过滤
    /// </summary>
    public interface IExporterHeaderFilter
    {
        /// <summary>
        /// 过滤列头（可以在此处理列名、是否隐藏等）
        /// </summary>
        /// <returns></returns>
        ExporterHeaderInfo Filter(ExporterHeaderInfo exporterHeaderInfo);
    }
}
