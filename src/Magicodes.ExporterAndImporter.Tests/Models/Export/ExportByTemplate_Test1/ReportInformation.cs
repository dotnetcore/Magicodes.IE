using System.Collections.Generic;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export.ExportByTemplate_Test1
{
    public class ReportInformation
    {
        public string CustomerName { get; set; }
        public string Date { get; set; }
        public string Contacts { get; set; }
        public string ContactsNumber { get; set; }
        public string SystemExhaustPressure { get; set; } = "NaN";
        public string SystemDewPressure { get; set; } = "NaN";
        public string SystemDayFlow { get; set; } = "NaN";

        public List<AirCompressor> AirCompressors { get; set; }
        public List<AfterProcessing> AfterProcessings { get; set; }
        public List<Suggest> Suggests { get; set; }
        public List<SystemPressureHisotry> SystemPressureHisotries { get; set; }
    }

}
