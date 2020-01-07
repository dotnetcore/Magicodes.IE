using System;
using System.Collections.Generic;
using System.Text;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    /// <summary>
    /// 教材订购信息
    /// </summary>
    public class TextbookOrderInfo
    {
        public string Company { get; }
        public string Address { get; }
        public string Contact { get; }
        public string Tel { get; }
        public string Watchmaker { get; }
        public string Time { get; }
        public List<BookInfo> BookInfo { get; }

        public TextbookOrderInfo(string company, string address, string contact, string tel, string watchmaker, string time, List<BookInfo> bookInfo)
        {
            Company = company;
            Address = address;
            Contact = contact;
            Tel = tel;
            Watchmaker = watchmaker;
            Time = time;
            BookInfo = bookInfo;
        }
    }
}
