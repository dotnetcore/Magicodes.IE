using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Magicodes.ExporterAndImporter.Core.Extension
{
    /// <summary>
    /// 
    /// </summary>
    public static class ReadManifestExtensions
    {
        /// <summary>
        /// 读取嵌入式资源
        /// </summary>
        /// <param name="assembly"></param>
        /// <param name="embeddedFileName"></param>
        /// <returns></returns>
        public static string ReadManifestString(this Assembly assembly, string embeddedFileName)
        {
            var resourceName = assembly.GetManifestResourceNames().First(s =>
                s.EndsWith(embeddedFileName, StringComparison.CurrentCultureIgnoreCase));

            using (var stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    throw new InvalidOperationException($"无法加载嵌入式资源，请确认路径是否正确：{embeddedFileName}。");
                }

                using (var reader = new StreamReader(stream, Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
        }
    }
}
