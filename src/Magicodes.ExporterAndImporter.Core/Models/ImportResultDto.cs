using System.Collections.Generic;

namespace Magicodes.ExporterAndImporter.Core.Models
{
    /// <summary>
    /// 
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ImportResultDto<T>
    {
        /// <summary>
        /// 
        /// </summary>
        public bool IsValid { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public List<T> Data { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public List<string> Errors { get; set; }
    }
}