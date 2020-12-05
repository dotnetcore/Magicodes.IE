// Copyright (c) .NET Core Community. All rights reserved.
// Licensed under the MIT License. See License in the project root for license information.

using Magicodes.ExporterAndImporter.Core;
using System;

namespace Magicodes.ExporterAndImporter.Excel
{

    [AttributeUsage(AttributeTargets.Class | AttributeTargets.Property)]
    public class ExcelImporterAttribute : ImporterAttribute
    {
        /// <summary>
        ///     指定Sheet名称(获取指定Sheet名称)
        ///     为空则自动获取第一个
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        ///     截止读取的列数（从1开始，如果已设置，则将支持空行以及特殊列）
        /// </summary>
        public int? EndColumnCount { get; set; }

        /// <summary>
        ///     是否标注错误（默认为true）
        /// </summary>
        public bool IsLabelingError { get; set; } = true;

        /// <summary>
        /// Sheet顶部导入描述
        /// </summary>
        public string ImportDescription { get; set; }

        /// <summary>
        /// Sheet顶部导入描述高度(换行可能无法自动设定高度,默认为Excel的默认行高)
        /// </summary>
        public double DescriptionHeight { get; set; } = 13.5;

        /// <summary>
        ///     是否仅导出错误数据
        /// </summary>
        public bool IsOnlyErrorRows { get; set; }
    }
}