﻿using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
        /// <param name="filePath"></param>
        /// <param name="lableFilePath"></param>
        /// <returns>返回一个字典，Key为Sheet名，Value为Sheet对应类型的object装箱，使用时做强转</returns>
        Task<Dictionary<string, ImportResult<object>>> ImportMultipleSheet<T>(string filePath,string lableFilePath) where T : class, new();
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
