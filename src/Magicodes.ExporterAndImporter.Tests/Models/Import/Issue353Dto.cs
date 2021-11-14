using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using System;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import
{
    [ExcelImporter(IsLabelingError = true)]
    public class Issue353Dto
    {
        [ImporterHeader(Name = "创建时间")]
        public DateTime CreatedOn { get; set; }

        [ImporterHeader(Name = "修改时间")]
        public DateTime ModifiedOn { get; set; }
    }
}
