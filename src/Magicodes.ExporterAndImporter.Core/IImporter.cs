// ======================================================================
// 
//           filename : IImporter.cs
//           description :
// 
//           created by 雪雁 at  2019-09-11 13:51
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

using Magicodes.ExporterAndImporter.Core.Models;
using System.IO;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Core
{
    /// <summary>
    ///     导入
    /// </summary>
    public interface IImporter
    {
        /// <summary> 
        ///     生成导入模板
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        Task<ExportFileInfo> GenerateTemplate<T>(string fileName) where T : class, new();

        /// <summary>
        ///     生成导入模板
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns>二进制字节</returns>
        Task<byte[]> GenerateTemplateBytes<T>() where T : class, new();

        /// <summary>
        /// 导入模型验证数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <param name="labelingFilePath">标注文件路径</param>
        /// <returns></returns>
        Task<ImportResult<T>> Import<T>(string filePath, string labelingFilePath = null) where T : class, new();

        /// <summary>
        /// 导入模型验证数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream">文件流</param>
        /// <returns></returns>
        Task<ImportResult<T>> Import<T>(Stream stream) where T : class, new();
    }
}