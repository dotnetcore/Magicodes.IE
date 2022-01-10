using System;
using System.ComponentModel.DataAnnotations;

namespace Magicodes.ExporterAndImporter.Tests.Models.Import.Issue305
{
    public class ProjectInfoImportDto
    {
        /// <summary>
        /// 项目编号
        /// </summary>
        [Display(Name = "项目编号"), Required]
        public string Code { get; set; }

        /// <summary>
        /// 项目名称
        /// </summary>
        [Display(Name = "项目名称"), Required]
        public string Name { get; set; }

        /// <summary>
        /// 合同金额
        /// </summary>
        [Display(Name = "合同金额"), Required]
        public decimal? ContractAmount { get; set; }

        /// <summary>
        /// 办卡数量
        /// </summary>
        [Display(Name = "办卡数量")]
        public string CardNumber { get; set; }

        /// <summary>
        /// 套餐月费(元)
        /// </summary>
        [Display(Name = "套餐月费(元)")]
        public string MonthlyFee { get; set; }

        /// <summary>
        /// 套餐包含流量(国内)
        /// </summary>
        [Display(Name = "套餐包含流量(国内)")]
        public string Flow { get; set; }

        /// <summary>
        /// 供应商
        /// </summary>
        [Display(Name = "供应商名称"), Required]
        public string Supplier { get; set; }

        /// <summary>
        /// 合同开始日期
        /// </summary>
        [Display(Name = "合同开始日期"), Required]
        public DateTime StartDate { get; set; }

        /// <summary>
        /// 合同截止日期
        /// </summary>
        [Display(Name = "合同截止日期"), Required]
        public DateTime EndDate { get; set; }

        /// <summary>
        /// 领取部门
        /// </summary>
        [Display(Name = "领取部门"), Required]
        public string Department { get; set; }

        /// <summary>
        /// 负责人
        /// </summary>
        [Display(Name = "负责人"), Required]
        public string Functionary { get; set; }

        /// <summary>
        /// 卡使用截止日期（项目移交给业主单位的时间）
        /// </summary>
        [Display(Name = "卡使用截止日期"), Required]
        public DateTime HandoverDate { get; set; }

        /// <summary>
        /// 合同号
        /// </summary>
        [Display(Name = "合同号"), Required]
        public string ContractNumber { get; set; }

        /// <summary>
        /// 缴费节点(月)
        /// </summary>
        [Display(Name = "缴费节点(月)"), Required, Range(1, 12, ErrorMessage = "缴费节点为无效值")]
        public int PaymentNode { get; set; }

        /// <summary>
        /// 种类
        /// </summary>
        [Display(Name = "种类"), Required]
        public string CategoryCode { get; set; }
    }
}
