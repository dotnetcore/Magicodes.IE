// ======================================================================
// 
//           filename : BookInfo.cs
//           description :
// 
//           created by 雪雁 at  -- 
//           文档官网：https://docs.xin-lai.com
//           公众号教程：麦扣聊技术
//           QQ群：85318032（编程交流）
//           Blog：http://www.cnblogs.com/codelove/
// 
// ======================================================================

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    /// <summary>
    /// 教材信息
    /// </summary>
    public class BookInfo
    {
        /// <summary>
        /// 行号
        /// </summary>
        public int RowNo { get; }

        /// <summary>
        /// 书号
        /// </summary>
        public string No { get; }

        /// <summary>
        /// 书名
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// 主编
        /// </summary>
        public string EditorInChief { get; }

        /// <summary>
        /// 出版社
        /// </summary>
        public string PublishingHouse { get; }

        /// <summary>
        /// 定价
        /// </summary>
        public string Price { get; }

        /// <summary>
        /// 采购数量
        /// </summary>
        public int PurchaseQuantity { get; }

        /// <summary>
        /// 备注
        /// </summary>
        public string Remark { get; }

        /// <summary>
        /// 封面
        /// </summary>
        public string Cover { get; set; }

        public BookInfo() { }

        public BookInfo(int rowNo, string no, string name, string editorInChief, string publishingHouse, string price, int purchaseQuantity, string remark)
        {
            RowNo = rowNo;
            No = no;
            Name = name;
            EditorInChief = editorInChief;
            PublishingHouse = publishingHouse;
            Price = price;
            PurchaseQuantity = purchaseQuantity;
            Remark = remark;
        }
    }
}