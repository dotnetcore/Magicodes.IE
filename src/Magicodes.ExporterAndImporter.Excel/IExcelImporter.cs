using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Models;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Magicodes.ExporterAndImporter.Excel
{
    /// <summary>
    /// Excel 导入程序
    /// </summary>
    public interface IExcelImporter : IImporter
    {
        /// <summary>
        /// 导出业务错误数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath">文件路径</param>
        /// <param name="bussinessErrorDataList">错误数据</param>
        /// <param name="msg">成功:错误数据返回路径,失败 返回错误原因</param>
        /// <returns></returns>
        bool OutputBussinessErrorData<T>(string filePath, List<DataRowErrorInfo> bussinessErrorDataList, out string msg) where T : class, new();
        
        /// <summary>
        /// 导出业务错误数据
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="stream">流</param>
        /// <param name="bussinessErrorDataList">错误数据</param>
        /// <param name="fileByte">成功:错误数据返回文件流字节,失败 返回null</param>
        /// <returns></returns>
        bool OutputBussinessErrorData<T>(Stream stream, List<DataRowErrorInfo> bussinessErrorDataList, out byte[] fileByte) where T : class, new();

        /// <summary>
        /// 导入多个Sheet数据
        /// </summary>
        /// <typeparam name="T">Excel类</typeparam>
        /// <param name="filePath"></param>
        /// <returns>返回一个字典，Key为Sheet名，Value为Sheet对应类型的object装箱，使用时做强转</returns>
        Task<Dictionary<string, ImportResult<object>>> ImportMultipleSheet<T>(string filePath) where T : class, new();
        
        /// <summary>
        /// 导入多个Sheet数据
        /// </summary>
        /// <typeparam name="T">Excel类</typeparam>
        /// <param name="stream"></param>
        /// <returns>返回一个字典，Key为Sheet名，Value为Sheet对应类型的object装箱，使用时做强转</returns>
        Task<Dictionary<string, ImportResult<object>>> ImportMultipleSheet<T>(Stream stream) where T : class, new();

        /// <summary>
        /// 导入多个相同类型的Sheet数据
        /// </summary>
        /// <typeparam name="T">Excel类</typeparam>
        /// <typeparam name="TSheet">Sheet类</typeparam>
        /// <param name="filePath"></param>
        /// <returns>返回一个字典，Key为Sheet名，Value为Sheet对应类型TSheet</returns>
        Task<Dictionary<string, ImportResult<TSheet>>> ImportSameSheets<T, TSheet>(string filePath)
            where T : class, new() where TSheet : class, new();
    }
}
