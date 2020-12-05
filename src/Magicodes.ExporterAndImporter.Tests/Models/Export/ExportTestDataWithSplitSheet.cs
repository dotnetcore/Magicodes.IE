// Copyright (c) .NET Core Community. All rights reserved.
// Licensed under the MIT License. See License in the project root for license information.

using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using System;
using OfficeOpenXml.Table;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    [ExcelExporter(Name = "测试2", TableStyle = TableStyles.None, AutoFitAllColumn = true, MaxRowNumberOnASheet = 100)]
    public class ExportTestDataWithSplitSheet
    {
        [ExporterHeader(DisplayName = "加粗文本", IsBold = true)]
        public string Text { get; set; }

        [ExporterHeader(DisplayName = "普通文本")] public string Text2 { get; set; }

        [ExporterHeader(DisplayName = "忽略", IsIgnore = true)]
        public string Text3 { get; set; }

        [ExporterHeader(DisplayName = "数值", Format = "#,##0")]
        public decimal Number { get; set; }

        [ExporterHeader(DisplayName = "名称", IsAutoFit = true)]
        public string Name { get; set; }

        /// <summary>
        /// 时间测试
        /// </summary>
        [ExporterHeader(DisplayName = "日期1", Format = "yyyy-MM-dd")]
        public DateTime Time1 { get; set; }

        /// <summary>
        /// 时间测试
        /// </summary>
        [ExporterHeader(DisplayName = "日期2", Format = "yyyy-MM-dd HH:mm:ss")]
        public DateTime? Time2 { get; set; }

        public DateTime Time3 { get; set; }

        public DateTime Time4 { get; set; }
    }
}