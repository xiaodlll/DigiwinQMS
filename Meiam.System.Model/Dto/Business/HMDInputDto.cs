using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Text;

namespace Meiam.System.Model.Dto {

    /// <summary>
    /// 恒铭达Dto
    /// </summary>
    public class HMDInputDto {

        [Display(Name = "生产工单")]
        public string MOID { get; set; }

        [Display(Name = "料号")]
        public string ITEMID { get; set; }

        [Display(Name = "生产机台号")]
        public string MA_DEVID { get; set; }

        [Display(Name = "生产日期")]
        public string PRO_DATE { get; set; }

        [Display(Name = "检测类型")]
        public string INSPECT_PUR { get; set; }

        [Display(Name = "检测机台号")]
        public string IN_DEVID { get; set; }

        [Display(Name = "检验人员")]
        public string PEOPLEID { get; set; }

        [Display(Name = "检测日期")]
        public string INSPECT_DATE { get; set; }

        [Display(Name = "检验明细")]
        public List<HMDInputDetailDto> Details { get; set; }
    }

    /// <summary>
    /// 恒铭达明细Dto
    /// </summary>
    public class HMDInputDetailDto {

        [Display(Name = "检验项目名称")]
        public string INSPECT_PROGRESSNAME { get; set; }

        [Display(Name = "标准值")]
        public string STD_VALUE { get; set; }

        [Display(Name = "上公差")]
        public string MAX_VALUE { get; set; }

        [Display(Name = "下公差")]
        public string MIN_VALUE { get; set; }

        [Display(Name = "上限值")]
        public string UP_VALUE { get; set; }

        [Display(Name = "下限值")]
        public string DOWN_VALUE { get; set; }

        [Display(Name = "样本值")]
        public string[] SAMPLE_VALUES { get; set; }

    }
}
