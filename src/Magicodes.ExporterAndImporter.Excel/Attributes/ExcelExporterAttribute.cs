// Copyright (c) .NET Core Community. All rights reserved.
// Licensed under the MIT License. See License in the project root for license information.

using Magicodes.ExporterAndImporter.Core;
using OfficeOpenXml.Table;

namespace Magicodes.ExporterAndImporter.Excel
{
    /// <summary>
    ///     Excel导出特性
    /// </summary>
    public class ExcelExporterAttribute : ExporterAttribute
    {
        /// <summary>
        ///  输出类型
        /// </summary>
        public ExcelOutputTypes ExcelOutputType { get; set; } = ExcelOutputTypes.DataTable;

        /// <summary>
        ///     自动居中(设置后为全局居中显示)
        /// </summary>
        public bool AutoCenter { get; set; }

        /// <summary>
        ///     表头位置
        /// </summary>
        public int HeaderRowIndex { get; set; } = 1;


        /// <summary>
        ///     表格样式风格
        /// </summary>
        public TableStyles TableStyle { get; set; } = TableStyles.None;
    }
}