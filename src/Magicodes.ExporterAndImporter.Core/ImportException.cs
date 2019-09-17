using System;
using System.Runtime.Serialization;

namespace Magicodes.ExporterAndImporter.Core
{
    /// <summary>
    /// 自定义导入Exception
    /// </summary>
    [Serializable]
    public class ImportException : Exception
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        public ImportException(string message) : base(message)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="message"></param>
        /// <param name="innerException"></param>
        public ImportException(string message, Exception innerException) : base(message, innerException)
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="info"></param>
        /// <param name="context"></param>
        protected ImportException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}