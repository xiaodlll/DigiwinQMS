using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace Meiam.System.Model.Dto {

    /// <summary>
    /// COC报告
    /// </summary>
    public class COCInputDto {

        /// <summary>
        /// 描述 : COCID 
        /// 空值 : False
        /// 默认 : 
        /// </summary>
        [Display(Name = "COCID")]
        public string COCID { get; set; }

        /// <summary>
        /// 描述 : VLOOKCODE 
        /// 空值 : False
        /// 默认 : 
        /// </summary>
        [Display(Name = "VLOOKCODE")]
        public string VLOOKCODE { get; set; }

        /// <summary>
        /// 描述 : FIX_VALUE 
        /// 空值 : False
        /// 默认 : 
        /// </summary>
        [Display(Name = "FIX_VALUE")]
        public string[] FIX_VALUE { get; set; }

        /// <summary>
        /// 描述 : ID 
        /// 空值 : False
        /// 默认 : 
        /// </summary>
        [Display(Name = "ID")]
        public string ID { get; set; }

        public string INSPECT_PROGRESSID { get; set; }
        
        public string INSPECT_DEV1ID { get; set; }

        public string DOCTYPE { get; set; }
        
    }

}
