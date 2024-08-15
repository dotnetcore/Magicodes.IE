using Microsoft.IO;
using System;
using System.IO;

namespace OfficeOpenXml.Utils
{
    public static class RecyclableMemoryStream
    {
       private static readonly Lazy<RecyclableMemoryStreamManager> recyclableMemoryStreamManager = new Lazy<RecyclableMemoryStreamManager>(() =>
       {
           var option = new RecyclableMemoryStreamManager.Options();
		   option.MaximumSmallPoolFreeBytes = 64 * 1024 * 1024;
           option.MaximumLargePoolFreeBytes = 64 * 1024 * 32;
           option.AggressiveBufferReturn = true;
		   return new RecyclableMemoryStreamManager(option);
       });
       private static RecyclableMemoryStreamManager RecyclableMemoryStreamManager
       {
           get
           {
               return recyclableMemoryStreamManager.Value;
           }
       }
        private const string TagSource = "Magicodes.EPPlus";

        internal static MemoryStream GetStream()
        {
            return RecyclableMemoryStreamManager.GetStream(TagSource);
        }

        internal static MemoryStream GetStream(byte[] array)
        {
            return RecyclableMemoryStreamManager.GetStream(array);
        }

        internal static MemoryStream GetStream(int capacity)
        {
            return RecyclableMemoryStreamManager.GetStream(TagSource, capacity);
        }
    }
}
