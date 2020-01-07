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
        public int RowNo { get; }
        public string No { get; }
        public string Name { get; }
        public string EditorInChief { get; }
        public string PublishingHouse { get; }
        public string Price { get; }
        public int PurchaseQuantity { get; }
        public string Remark { get; }

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