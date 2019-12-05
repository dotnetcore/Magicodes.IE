// ======================================================================
// 
//           filename : TestBase.cs
//           description :
// 
//           created by 雪雁 at  2019-10-18 14:07
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Magicodes.ExporterAndImporter.Tests
{
    public class TestBase
    {
        public TestBase()
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            //Encoding.UTF8
            Encoding.GetEncoding(65001);
        }
        /// <summary>
        ///     获取根目录
        /// </summary>
        /// <returns></returns>
        public string GetTestRootPath()
        {
            return Directory.GetCurrentDirectory();
        }

        /// <summary>
        ///     获取测试文件路径
        /// </summary>
        /// <param name="paths"></param>
        /// <returns></returns>
        public string GetTestFilePath(params string[] paths)
        {
            var rootPath = GetTestRootPath();
            var list = new List<string>
            {
                rootPath
            };
            list.AddRange(paths);
            return Path.Combine(list.ToArray());
        }

        /// <summary>
        ///     删除文件
        /// </summary>
        public void DeleteFile(string path)
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}