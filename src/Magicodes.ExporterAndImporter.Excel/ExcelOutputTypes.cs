// ======================================================================
// 
//           filename : ExcelExporterAttribute.cs
//           description :
// 
//           created by 雪雁 at  2020-03-25 13:51
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

namespace Magicodes.ExporterAndImporter.Excel
{
    /// <summary>
    /// 输出类型
    /// </summary>
    public enum ExcelOutputTypes
    {
        /// <summary>
        /// Excel数据表格
        /// </summary>
        DataTable = 0,

        /// <summary>
        /// 普通的单元格写入
        /// </summary>
        None = 1
    }
}