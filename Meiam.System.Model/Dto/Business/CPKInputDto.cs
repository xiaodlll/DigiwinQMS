using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace Meiam.System.Model.Dto {

    /// <summary>
    /// 拉力机检测报告
    /// </summary>
    public class CPKInputDto {

        /// <summary>
        /// 描述 : 主档ID 
        /// 空值 : False
        /// 默认 : 
        /// </summary>
        [Display(Name = "主档ID")]
        public string INSPECT_DEV2ID { get; set; }

        /// <summary>
        /// 描述 : 用户名 
        /// 空值 : False
        /// 默认 : 
        /// </summary>
        [Display(Name = "用户名")]
        public string UserName { get; set; }
    }

    /// <summary>
    /// 批量拉力机检测报告
    /// </summary>
    public class CPKInputByCodeDto
    {

        /// <summary>
        /// 描述 : DOC_CODE 
        /// 空值 : False
        /// 默认 : 
        /// </summary>
        [Display(Name = "DOC_CODE")]
        public string DOC_CODE { get; set; }


        /// <summary>
        /// 描述 : INSPECT_DEV 
        /// 空值 : False
        /// 默认 : 
        /// </summary>
        [Display(Name = "INSPECT_DEV")]
        public string INSPECT_DEV { get; set; }

        /// <summary>
        /// 描述 : 用户名 
        /// 空值 : False
        /// 默认 : 
        /// </summary>
        [Display(Name = "用户名")]
        public string UserName { get; set; }
    }
}
