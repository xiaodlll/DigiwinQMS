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
        /// 描述 : 检验单号 
        /// 空值 : False
        /// 默认 : 
        /// </summary>
        [Display(Name = "用户名")]
        public string UserName { get; set; }
    }
}
