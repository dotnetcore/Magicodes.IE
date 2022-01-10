using System;
using System.Collections.Generic;
using System.Text;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    public class Issue296
    {
        public string ReportTitle { get; set; }
        public string BeginDate { get; set; }
        public string EndDate { get; set; }
        public 播放大厅营收报表[] 播放大厅营收报表 { get; set; }
        public 播放大厅能耗情况[] 播放大厅能耗情况 { get; set; }
        public 安全情况[] 安全情况 { get; set; }
        public 考勤情况[] 考勤情况 { get; set; }
    }

    public class 播放大厅营收报表
    {
        public string EquipName { get; set; }
        public string 放映场次 { get; set; }
        public int 取消场次 { get; set; }
        public string 售票数量 { get; set; }
        public string 入场人数 { get; set; }
        public string 入场异常 { get; set; }
    }

    public class 播放大厅能耗情况
    {
        public string EquipName { get; set; }
        public string 放映设备 { get; set; }
        public int 放映空调 { get; set; }
        public string _4D设备 { get; set; }
        public string 能耗异常 { get; set; }
        public string 冷凝机组 { get; set; }
        public string 售卖区 { get; set; }
    }

    public class 安全情况
    {
        public string EquipName { get; set; }
        public string 时间 { get; set; }
        public string 位置 { get; set; }
        public string 次数 { get; set; }
    }

    public class 考勤情况
    {
        public string EquipName { get; set; }
        public string 出勤 { get; set; }
        public string 休假 { get; set; }
        public string 迟到 { get; set; }
        public string 缺勤 { get; set; }
        public string 总人数 { get; set; }
    }
}