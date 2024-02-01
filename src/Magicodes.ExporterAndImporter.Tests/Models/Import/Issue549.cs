using Magicodes.ExporterAndImporter.Core;

namespace Magicodes.IE.Tests.Models.Import;

public class Issue549
{
    [ImporterHeader(Name = "姓名", AutoTrim = true)]
    public string Name { get; set; }

    [ImporterHeader(Name = "年龄", AutoTrim = true)]
    public decimal Age { get; set; }

    [ImporterHeader(Name = "身高", AutoTrim = true)]
    public decimal Height { get; set; }
}
