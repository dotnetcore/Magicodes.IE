using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using System;
using System.ComponentModel.DataAnnotations;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    [ExcelImporter(IsLabelingError = true)]
    [ExcelExporter(Name = "AOI内层", TableStyle = OfficeOpenXml.Table.TableStyles.Medium4)]
    [Serializable]
    public class Issue179
    {
        [ImporterHeader(IsIgnore = true)]
        [ExporterHeader(IsIgnore = true)]
        public long Id
        {
            get; set;
        }
        ///

        /// 用于页面显示的唯一ID
        ///
        [ImporterHeader(IsIgnore = true)]
        [ExporterHeader(IsIgnore = true)]
        public string UniqueId { get; set; }
        ///

        /// 用于页面使用，是否存在子节点
        ///
        [ImporterHeader(IsIgnore = true)]
        [ExporterHeader(IsIgnore = true)]
        public bool HasChildren { get; set; }
        ///

        /// 日期
        ///
        [ImporterHeader(Name = "日期")]
        [Required(ErrorMessage = "日期不能为空")]
        [ExporterHeader(DisplayName = "日期", Format = "yyyy-m-d")]
        public DateTime CheckDate { get; set; }
        ///

        /// 员工姓名
        ///
        [ImporterHeader(Name = "员工姓名")]
        [Required(ErrorMessage = "员工姓名不能为空")]
        [MaxLength(20, ErrorMessage = "测试治具最大长度20")]
        [ExporterHeader(DisplayName = "员工姓名", ColumnIndex = 0)]
        public string CheckUserName { get; set; }

        [ImporterHeader(Name = "料号")]
        [Required(ErrorMessage = "料号不能为空")]
        [ExporterHeader(DisplayName = "料号", ColumnIndex = 1)]
        public string InvPartNumber { get; set; }
        /// <summary>
        /// 检查总数PANEL
        /// </summary>
        [ImporterHeader(Name = "检测总数(PNL)")]
        [Required(ErrorMessage = "检测总数(PNL)不能为空")]
        [ExporterHeader(DisplayName = "检测总数(PNL)")]
        public int CheckQTY_PNL { get; set; }
        /// <summary>
        /// 排版SET数
        /// </summary>
        [ImporterHeader(Name = "排版SET数")]
        [Required(ErrorMessage = "排版SET数不能为空")]
        [ExporterHeader(DisplayName = "排版SET数")]
        public int PnlSet { get; set; }
        /// <summary>
        /// 排版SET总数=CheckQTY_PNL* PnlSet
        /// </summary>
        [ImporterHeader(IsIgnore = true)]
        [ExporterHeader(DisplayName = "排版SET总数")]
        public int CheckQTY_SET { get; set; }

        [ImporterHeader(Name = "PO")]
        [MaxLength(40, ErrorMessage = "测试治具最大长度40")]
        [ExporterHeader(DisplayName = "PO")]
        public string PONumber { get; set; }

        //层别信息--------------------------
        /// <summary>
        /// 层别名称
        /// </summary>
        [ImporterHeader(Name = "层别")]
        [Required(ErrorMessage = "层别不能为空")]
        [MaxLength(10, ErrorMessage = "层别最大长度10")]
        [ExporterHeader(DisplayName = "层别")]
        public string LayerName { get; set; }
        /// <summary>
        /// 不良数=修补，报废所有缺陷数总和
        /// </summary>
        [ImporterHeader(IsIgnore = true)]
        [ExporterHeader(DisplayName = "不良数(点)")]
        public int BadQTY { get; set; }
        /// <summary>
        /// 报废数=报废数总和
        /// </summary>
        [ImporterHeader(IsIgnore = true)]
        [ExporterHeader(DisplayName = "报废总数(SET)")]
        public int ScrapQTY { get; set; }
        /// <summary>
        /// 一次直行率(点)=(CheckQTY_SET*2-BadQTY)/(CheckQTY_SET*2)
        /// </summary>
        [ImporterHeader(IsIgnore = true)]
        [ExporterHeader(DisplayName = "一次直行率(点)", Format = "0.00%")]
        public decimal? DirectRate { get; set; }
        /// <summary>
        /// 修补后良率=(CheckQTY_SET-ScrapQTY)/CheckQTY_SET
        /// </summary>
        [ImporterHeader(IsIgnore = true)]
        [ExporterHeader(DisplayName = "修补后良率(点)", Format = "0.00%")]
        public decimal? RepairRate { get; set; }

        //具体项目
        /// <summary>
        /// 判定项  - 修补/报废
        /// </summary>
        [ImporterHeader(Name = "判定项")]
        [Required(ErrorMessage = "判定项不能为空")]
        [ValueMapping("修补", "修补")]
        [ValueMapping("报废", "报废")]
        [ExporterHeader(DisplayName = "判定项")]
        public string JudgeItem { get; set; }

        [ImporterHeader(Name = "开路")]
        [ExporterHeader(DisplayName = "开路")]
        public int? DefectQty1 { get; set; }

        [ImporterHeader(Name = "缺口")]
        [ExporterHeader(DisplayName = "缺口")]
        public int? DefectQty2 { get; set; }

        [ImporterHeader(Name = "短路")]
        [ExporterHeader(DisplayName = "短路")]
        public int? DefectQty3 { get; set; }

        [ImporterHeader(Name = "曝光不良")]
        [ExporterHeader(DisplayName = "曝光不良")]
        public int? DefectQty4 { get; set; }

        [ImporterHeader(Name = "蚀刻不净")]
        [ExporterHeader(DisplayName = "蚀刻不净")]
        public int? DefectQty5 { get; set; }

        [ImporterHeader(Name = "残铜毛刺")]
        [ExporterHeader(DisplayName = "残铜毛刺")]
        public int? DefectQty6 { get; set; }

        [ImporterHeader(Name = "孔内无铜")]
        [ExporterHeader(DisplayName = "孔内无铜")]
        public int? DefectQty7 { get; set; }

        [ImporterHeader(Name = "孔内毛刺")]
        [ExporterHeader(DisplayName = "孔内毛刺")]
        public int? DefectQty8 { get; set; }

        [ImporterHeader(Name = "掉膜")]
        [ExporterHeader(DisplayName = "掉膜")]
        public int? DefectQty9 { get; set; }

        [ImporterHeader(Name = "线细")]
        [ExporterHeader(DisplayName = "线细")]
        public int? DefectQty10 { get; set; }

        [ImporterHeader(Name = "铜渣")]
        [ExporterHeader(DisplayName = "铜渣")]
        public int? DefectQty11 { get; set; }

        [ImporterHeader(Name = "针孔")]
        [ExporterHeader(DisplayName = "针孔")]
        public int? DefectQty12 { get; set; }

        [ImporterHeader(Name = "撞断线")]
        [ExporterHeader(DisplayName = "撞断线")]
        public int? DefectQty13 { get; set; }

        [ImporterHeader(Name = "折板")]
        [ExporterHeader(DisplayName = "折板")]
        public int? DefectQty14 { get; set; }

        [ImporterHeader(Name = "孔塞")]
        [ExporterHeader(DisplayName = "孔塞")]
        public int? DefectQty15 { get; set; }

        [ImporterHeader(Name = "孔偏破")]
        [ExporterHeader(DisplayName = "孔偏破")]
        public int? DefectQty16 { get; set; }

        [ImporterHeader(Name = "多/少孔")]
        [ExporterHeader(DisplayName = "多/少孔")]
        public int? DefectQty17 { get; set; }

        [ImporterHeader(Name = "压痕")]
        [ExporterHeader(DisplayName = "压痕")]
        public int? DefectQty18 { get; set; }

        [ImporterHeader(Name = "杂物")]
        [ExporterHeader(DisplayName = "杂物")]
        public int? DefectQty19 { get; set; }

        [ImporterHeader(Name = "划伤")]
        [ExporterHeader(DisplayName = "划伤")]
        public int? DefectQty20 { get; set; }

        [ImporterHeader(Name = "其它")]
        [ExporterHeader(DisplayName = "其它")]
        public int? DefectQty21 { get; set; }
    }
}
