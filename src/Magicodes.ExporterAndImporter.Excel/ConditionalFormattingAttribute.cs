using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Magicodes.IE.Excel
{
    public class ValueConditionalFormattingAttribute: Attribute
    {
        public string Value { get; set; }
    }
}
