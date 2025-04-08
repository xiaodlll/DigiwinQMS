using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace Meiam.System.Model.Dto
{

    /// <summary>
    /// 替换FTIR的PDF图片
    /// </summary>
    public class FTIRInputDto
    {

        /// <summary>
        /// 描述 : 检验单号 
        /// 空值 : False
        /// 默认 : 
        /// </summary>
        [Display(Name = "检验单号")]
        public string INSPECTCODE { get; set; }

        /// <summary>
        /// 描述 : 检验单类型
        /// 空值 : False
        /// 默认 : 
        /// </summary>
        [Display(Name = "检验单类型")]
        public string INSPECTTYPE { get; set; }

        /// <summary>
        /// 描述 : 用户名 
        /// 空值 : False
        /// 默认 : 
        /// </summary>
        [Display(Name = "用户名")]
        public string UserName { get; set; }
    }
}
