using System.Collections.Generic;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    /// <summary>
    /// 教材订购信息
    /// </summary>
    public class TextbookOrderInfo
    {
        /// <summary>
        /// 公司名称
        /// </summary>
        public string Company { get; }

        /// <summary>
        /// 地址
        /// </summary>
        public string Address { get; }

        /// <summary>
        /// 联系人
        /// </summary>
        public string Contact { get; }

        /// <summary>
        /// 电话
        /// </summary>
        public string Tel { get; }

        /// <summary>
        /// 制表人
        /// </summary>
        public string Watchmaker { get; }

        /// <summary>
        /// 时间
        /// </summary>
        public string Time { get; }

        public string ImageUrl { get; set; }

        /// <summary>
        /// 教材信息列表
        /// </summary>
        public List<BookInfo> BookInfos { get; }

        public TextbookOrderInfo(string company, string address, string contact, string tel, string watchmaker, string time, string imageUrl, List<BookInfo> bookInfo)
        {
            Company = company;
            Address = address;
            Contact = contact;
            Tel = tel;
            Watchmaker = watchmaker;
            Time = time;
            ImageUrl = imageUrl;
            BookInfos = bookInfo;
        }
    }
}
