using System;
using System.Collections.Generic;
using System.Text;

namespace Magicodes.ExporterAndImporter.Core.Models
{
    public class ImportResultDto<T>
    {
        public bool IsValid { get; set; }

        public List<T> Data { get; set; }

        public List<string> Errors { get; set; }
    }
}
