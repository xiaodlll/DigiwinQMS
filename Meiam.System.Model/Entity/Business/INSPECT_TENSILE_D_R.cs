﻿//------------------------------------------------------------------------------
// <auto-generated>
//     此代码已从模板生成手动更改此文件可能导致应用程序出现意外的行为。
//     如果重新生成代码，将覆盖对此文件的手动更改。
//     author MEIAM
// </auto-generated>
//------------------------------------------------------------------------------

using System.ComponentModel.DataAnnotations;
using System;
using System.Linq;
using System.Text;
using SqlSugar;


namespace Meiam.System.Model {
    ///<summary>
    ///拉力机样本结果档
    ///</summary>
    [SugarTable("INSPECT_TENSILE_D_R")]
    public class INSPECT_TENSILE_D_R {
        public INSPECT_TENSILE_D_R() {
        }

        /// <summary>
        /// 描述 : INSPECT_TENSILE_D_RID 
        /// 空值 : False
        /// 默认 : 
        /// </summary>
        [Display(Name = "INSPECT_TENSILE_D_RID")]
        [SugarColumn(IsPrimaryKey = true)]
        public string INSPECT_TENSILE_D_RID { get; set; }

        /// <summary>
        /// 描述 : INSPECT_TENSILE_DID
        /// 空值 : False
        /// 默认 : 
        /// </summary>
        [Display(Name = "INSPECT_TENSILE_DID")]
        public string INSPECT_TENSILE_DID { get; set; }

        /// <summary>
        /// 描述 : INSPECT_DEV1ID 
        /// 空值 : 可根据实际情况设置
        /// 默认 : 
        /// </summary>
        [Display(Name = "INSPECT_DEV1ID")]
        public string INSPECT_DEV1ID { get; set; }

        /// <summary>
        /// 描述 : SAMPLEID
        /// 空值 : False
        /// 默认 : 
        /// </summary>
        [Display(Name = "SAMPLEID")]
        public string SAMPLEID { get; set; }

        /// <summary>
        /// 描述 : MaxValue 
        /// 空值 : 可根据实际情况设置
        /// 默认 : 
        /// </summary>
        [Display(Name = "MaxValue")]
        public decimal MaxValue { get; set; }

        /// <summary>
        /// 描述 : MinValue 
        /// 空值 : 可根据实际情况设置
        /// 默认 : 
        /// </summary>
        [Display(Name = "MinValue")]
        public decimal MinValue { get; set; }

        /// <summary>
        /// 描述 : AvgValue 
        /// 空值 : 可根据实际情况设置
        /// 默认 : 
        /// </summary>
        [Display(Name = "AvgValue")]
        public decimal AvgValue { get; set; }

    }
}