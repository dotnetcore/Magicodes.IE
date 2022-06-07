using Magicodes.ExporterAndImporter.Core;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Magicodes.IE.Tests.Models.Import
{
    public class Issue377
    {
        [ImporterHeader(Name = "公司编号", AutoTrim = true)]
        public string NO { get; set; }

        [ImporterHeader(Name = "销售价", AutoTrim = true)]
        public decimal Price { get; set; }

    }
}
