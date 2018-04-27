using System;
using System.Runtime.Serialization;

namespace Magicodes.ExporterAndImporter.Core
{
    [Serializable]
    public class ImportException : Exception
    {
        public ImportException()
        {
        }

        public ImportException(string message) : base(message)
        {
        }

        public ImportException(string message, Exception innerException) : base(message, innerException)
        {
        }

        protected ImportException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}