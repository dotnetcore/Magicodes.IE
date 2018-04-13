using System;
using System.Collections.Generic;
using System.Text;

namespace Magicodes.ExporterAndImporter.Core.Extension
{
    public static class Extension
    {
        public static bool IsNullOrWhiteSpace(this string str)
        {
            return string.IsNullOrWhiteSpace(str);
        }
    }
}
