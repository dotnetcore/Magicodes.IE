// ======================================================================
// 
//           filename : ReadManifestExtensions.cs
//           description :
// 
//           created by 雪雁 at  2019-10-12 11:12
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Magicodes.ExporterAndImporter.Core.Extension
{
    /// <summary>
    /// </summary>
    public static class ReadManifestExtensions
    {
        /// <summary>
        ///     读取嵌入式资源
        /// </summary>
        /// <param name="assembly"></param>
        /// <param name="embeddedFileName"></param>
        /// <returns></returns>
        public static string ReadManifestString(this Assembly assembly, string embeddedFileName)
        {
            var resourceName = assembly.GetManifestResourceNames().First(s =>
                s.EndsWith(embeddedFileName, StringComparison.CurrentCultureIgnoreCase));

            using (var stream = assembly.GetManifestResourceStream(resourceName))
            using (var reader = new StreamReader(stream, Encoding.UTF8))
            {
                return reader.ReadToEnd();
            }

        }
    }
}