using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using OfficeOpenXml.Table;

namespace Magicodes.ExporterAndImporter.Tests.Models.Export
{
    /// <summary>
    /// 管线导出模型
    /// </summary>
    [ExcelExporter(Name = "管线模型", TableStyle = TableStyles.Light10)]
    public class GalleryLineExportModel
    {
        /// <summary>
        /// 介质名称
        /// </summary>
        [ExporterHeader(DisplayName = "介质名称")]
        public string GalleryMediumName { get; set; }

        /// <summary>
        /// 介质编码
        /// </summary>
        [ExporterHeader(DisplayName = "介质编码")]
        public string GalleryMediumCode { get; set; }

        ///<summary>
        ///管线编号
        ///</summary>
        [ExporterHeader(DisplayName = "管线编号")]
        public string LineNumber { get; set; } = string.Empty;

        ///<summary>
        ///管道规格(直径)
        ///</summary>
        [ExporterHeader(DisplayName = "管道规格(直径)")]
        public string Diameter { get; set; } = string.Empty;

        ///<summary>
        ///压力
        ///</summary>
        [ExporterHeader(DisplayName = "压力")]
        public string Pressure { get; set; } = string.Empty;

        ///<summary>
        ///管道级别
        ///</summary>
        [ExporterHeader(DisplayName = "管道级别")]
        public string Level { get; set; } = string.Empty;

        ///<summary>
        ///操作温度
        ///</summary>
        [ExporterHeader(DisplayName = "操作温度")]
        public string Temperature { get; set; } = string.Empty;

        ///<summary>
        ///管道材质
        ///</summary>
        [ExporterHeader(DisplayName = "管道材质")]
        public string PipeMaterial { get; set; } = string.Empty;

        ///<summary>
        ///自何处(来源)
        ///</summary>
        [ExporterHeader(DisplayName = "自何处(来源)")]
        public string WhereFrom { get; set; } = string.Empty;

        ///<summary>
        ///到何处(去向)
        ///</summary>
        [ExporterHeader(DisplayName = "到何处(去向)")]
        public string WhereTo { get; set; } = string.Empty;

        ///<summary>
        ///设计压力
        ///</summary>
        [ExporterHeader(DisplayName = "设计压力")]
        public string DesignPressure { get; set; } = string.Empty;

        ///<summary>
        ///设计温度
        ///</summary>
        [ExporterHeader(DisplayName = "设计温度")]
        public string DesignTemperature { get; set; } = string.Empty;

        ///<summary>
        ///管道设计(kg/h)
        ///</summary>
        [ExporterHeader(DisplayName = "管道设计(kg/h)")]
        public string KG { get; set; } = string.Empty;

        ///<summary>
        ///试验压力
        ///</summary>
        [ExporterHeader(DisplayName = "试验压力")]
        public string TestPressure { get; set; } = string.Empty;

        ///<summary>
        ///试验介质
        ///</summary>
        [ExporterHeader(DisplayName = "试验介质")]
        public string TestMedium { get; set; } = string.Empty;

        ///<summary>
        ///管道设计(m3/h)
        ///</summary>
        [ExporterHeader(DisplayName = "管道设计(m3/h)")]
        public string M3 { get; set; } = string.Empty;

        ///<summary>
        ///输送特性
        ///</summary>
        [ExporterHeader(DisplayName = "输送特性")]
        public string Transport { get; set; } = string.Empty;

        ///<summary>
        ///焊缝等级
        ///</summary>
        [ExporterHeader(DisplayName = "焊缝等级")]
        public string WeldGrade { get; set; } = string.Empty;

        ///<summary>
        ///探伤比例
        ///</summary>
        [ExporterHeader(DisplayName = "探伤比例")]
        public string FlawDetection { get; set; } = string.Empty;

        ///<summary>
        ///绝热类型
        ///</summary>
        [ExporterHeader(DisplayName = "绝热类型")]
        public string AdiabatType { get; set; } = string.Empty;

        ///<summary>
        ///绝热厚度
        ///</summary>
        [ExporterHeader(DisplayName = "绝热厚度")]
        public string AdiabatThickness { get; set; } = string.Empty;

        ///<summary>
        ///涂漆代号
        ///</summary>
        [ExporterHeader(DisplayName = "涂漆代号")]
        public string PaintCode { get; set; } = string.Empty;

        ///<summary>
        ///管道压力等级
        ///</summary> 
        [ExporterHeader(DisplayName = "管道压力等级")]
        public string PressurePipeLevel { get; set; } = string.Empty;

        ///<summary>
        ///备注
        ///</summary>
        [ExporterHeader(DisplayName = "备注")]
        public string Remark { get; set; } = string.Empty;

        ///<summary>
        ///修改标记
        ///</summary>
        [ExporterHeader(DisplayName = "修改标记")]
        public string UpdateMark { get; set; } = string.Empty;

    }
}
