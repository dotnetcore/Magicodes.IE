using Magicodes.ExporterAndImporter.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Excel
{
    /// <summary>
    /// Excel导出程序
    /// </summary>
    public interface IExcelExporter : IExporter, IExportFileByTemplate
    {
        //Excel独立API将添加到这里
    }
}
