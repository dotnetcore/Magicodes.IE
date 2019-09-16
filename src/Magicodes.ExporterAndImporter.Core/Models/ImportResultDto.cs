using System.Collections.Generic;

namespace Magicodes.ExporterAndImporter.Core.Models
{
    public class ImportResultDto<T>
    {
        public bool IsValid { get; set; }

        public List<T> Data { get; set; }

        public List<string> Errors { get; set; }
    }
}