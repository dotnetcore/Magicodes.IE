using System.Collections.Generic;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    public class ExportTestDataWithSingleColTpl
    {
        public List<ExportTestDataWithSingleCol> List { get; set; }
    }

    public class ExportTestDataWithSingleCol
    {
        public string Name { get; set; }
    }
}
