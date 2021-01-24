
/* 项目“Magicodes.ExporterAndImporter.Tests (netcoreapp3.1)”的未合并的更改
在此之前:
using System;
using System.Collections.Generic;
using System.Text;
using Magicodes.ExporterAndImporter.Core;
在此之后:
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel;
using System;
*/

/* 项目“Magicodes.ExporterAndImporter.Tests (net461)”的未合并的更改
在此之前:
using System;
using System.Collections.Generic;
using System.Text;
using Magicodes.ExporterAndImporter.Core;
在此之后:
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel;
using System;
*/

/* 项目“Magicodes.ExporterAndImporter.Tests (net5.0)”的未合并的更改
在此之前:
using System;
using System.Collections.Generic;
using System.Text;
using Magicodes.ExporterAndImporter.Core;
在此之后:
using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Core.Models;
using Magicodes.ExporterAndImporter.Excel;
using System;
*/
using Magicodes.ExporterAndImporter.Core;
using System.Collections.Generic;
using System.Text;
using System;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import
{
    [ExcelImporter(IsLabelingError = true, HeaderRowIndex = 2)]
    public class ImportPictureHeaderNotFirstRowDto
    {
        [ImporterHeader(Name = "加粗文本")]
        public string Text { get; set; }
        [ImporterHeader(Name = "普通文本")]
        public string Text2 { get; set; }

        /// <summary>
        /// 将图片写入到临时目录
        /// </summary>
        [ImportImageField(ImportImageTo = ImportImageTo.TempFolder)]
        [ImporterHeader(Name = "图1")]
        public string Img1 { get; set; }
        [ImporterHeader(Name = "数值")]
        public string Number { get; set; }
        [ImporterHeader(Name = "名称")]
        public string Name { get; set; }
        [ImporterHeader(Name = "日期")]
        public DateTime Time { get; set; }

        /// <summary>
        /// 将图片写入到临时目录
        /// </summary>
        [ImportImageField(ImportImageTo = ImportImageTo.TempFolder)]
        [ImporterHeader(Name = "图")]
        public string Img { get; set; }
    }
}
