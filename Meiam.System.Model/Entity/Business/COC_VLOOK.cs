using SqlSugar;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Meiam.System.Model {

    ///<summary>
    ///自定义数据源表
    ///</summary>
    [SugarTable("COC_VLOOK")]
    public class COC_VLOOK {
        public COC_VLOOK() {
        }

        /// <summary>
        /// 描述 : 自定义数据源ID 
        /// 空值 : False
        /// 默认 : 
        /// </summary>
        [Display(Name = "COC_VLOOKID")]
        [SugarColumn(IsPrimaryKey = true)]
        public string COC_VLOOKID { get; set; }

        /// <summary>
        /// 描述 : 通用数据源ID 
        /// 空值 : 可根据实际情况设置
        /// 默认 : 
        /// </summary>
        [Display(Name = "COULM002ID")]
        public string COULM002ID { get; set; }

        /// <summary>
        /// 描述 : 分组方式 
        /// 空值 : GOUPBY_001 分组合并   （分组字段显示，列分组字段隐藏）GOUPBY_002 全部合并   （分组字段，列分组字段隐藏）GOUPBY_003 行转列合并  （分组字段，列分组字段显示）
        /// 默认 : 
        /// </summary>
        [Display(Name = "GOUPBY")]
        public string GOUPBY { get; set; }

        /// <summary>
        /// 描述 : 分组字段 
        /// 空值 : 填写的内容需和COLUM001_COC中相同COLUM002ID的字段描述（COLUM001NAME）一致，多个分组字段用分号区分。例如：统计维度1；统计维度2
        /// 默认 : 
        /// </summary>
        [Display(Name = "GROUPBYNAME")]
        public string GROUPBYNAME { get; set; }

        /// <summary>
        /// 描述 : 列分组字段 
        /// 空值 : 必须包含在分组字段中，暂时只支持单个字段
        /// 默认 : 
        /// </summary>
        [Display(Name = "GROUPBYNAME_C")]
        public string GROUPBYNAME_C { get; set; }

        /// <summary>
        /// 描述 : 是否过滤固定查询条件 
        /// 空值 : 设定为1时:1.在执行数据源时增加检验单号的WHERE过滤条件2. WHERE条件的字段名由COULM002_COC.FIX_FILED获得;  3. 组成SQL表达式的符号是 in  (@FIX_VALUE)
        /// 默认 : 
        /// </summary>
        [Display(Name = "FIX_FILED")]
        public string FIX_FILED { get; set; }

    }

}
