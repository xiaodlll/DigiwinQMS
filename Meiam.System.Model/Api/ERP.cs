using System;

/// <summary>
/// Summary description for Class1
/// </summary>
namespace Meiam.System.Model
{
    /// <summary>
    /// ERP收料通知单
    /// </summary>
    public class LotNoticeRequest
    {
        [Required(ErrorMessage = "项次内码(ID)是必填字段")]
        public string ID { get; set; }

        [Required(ErrorMessage = "分录内码(ENTRYID)是必填字段")]
        public string ENTRYID { get; set; }

        [Required(ErrorMessage = "收料通知单号(ERP_ARRIVEDID)是必填字段")]
        public string ERP_ARRIVEDID { get; set; }

        [Required(ErrorMessage = "品号(ITEMID)是必填字段")]
        public string ITEMID { get; set; }

        [Required(ErrorMessage = "品名(ITEMNAME)是必填字段")]
        public string ITEMNAME { get; set; }

        [Required(ErrorMessage = "组织内码(ORGID)是必填字段")]
        public string ORGID { get; set; }

        [Required(ErrorMessage = "到货数量(LOT_QTY)是必填字段")]
        [Range(0, double.MaxValue, ErrorMessage = "到货数量必须是正数")]
        public decimal LOT_QTY { get; set; }

        [Required(ErrorMessage = "到货日期(APPLY_DATE)是必填字段")]
        [RegularExpression(@"^\d{4}-\d{2}-\d{2}$", ErrorMessage = "日期格式必须是YYYY-MM-DD")]
        public string APPLY_DATE { get; set; }

        // 规格型号
        public string MODEL_SPEC { get; set; }

        // 批号（检验员维护）
        public string LOTNO { get; set; }

        [Range(0, double.MaxValue, ErrorMessage = "长度必须是正数")]
        // 长度（单位：毫米）
        public decimal? LENGTH { get; set; }

        [Range(0, double.MaxValue, ErrorMessage = "宽度必须是正数")]
        // 宽度（单位：毫米）
        public decimal? WIDTH { get; set; }

        [Range(0, int.MaxValue, ErrorMessage = "卷数必须是正整数")]
        // 卷数
        public int? INUM { get; set; }

        [RegularExpression(@"^\d{4}-\d{2}-\d{2}$", ErrorMessage = "生产日期格式必须是YYYY-MM-DD")]
        // 生产日期
        public string PRO_DATE { get; set; }

        [RegularExpression(@"^\d{4}-\d{2}-\d{2}$", ErrorMessage = "失效日期格式必须是YYYY-MM-DD")]
        // 失效日期
        public string QUA_DATE { get; set; }

        // 供应商名称
        public string SUPPNAME { get; set; }
    }

    public class ApiResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; }
    }

    public class WorkOrderSyncRequest
    {
        [Required(ErrorMessage = "工单号（MOID）是必填字段")]
        public string MOID { get; set; }

        [Required(ErrorMessage = "内码（ID）是必填字段")]
        public string ID { get; set; }

        [Required(ErrorMessage = "品号（ITEMID）是必填字段")]
        public string ITEMID { get; set; }

        [Required(ErrorMessage = "品名（ITEMNAME）是必填字段")]
        public string ITEMNAME { get; set; }

        [Required(ErrorMessage = "组织内码ID（ORGID）是必填字段")]
        public string ORGID { get; set; }

        [Required(ErrorMessage = "建单时间（CREATEDATE）是必填字段")]
        [RegularExpression(@"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$",
            ErrorMessage = "建单时间的格式必须是YYYY-MM-DD HH:MM:SS")]
        public string CREATEDATE { get; set; }
    }

    public class LotCheckResultRequest
    {
        [Required(ErrorMessage = "品号（ITEMID）是必填字段")]
        public string ITEMID { get; set; }

        [Required(ErrorMessage = "检验日期（CHECKDATE）是必填字段")]
        [RegularExpression(@"^\d{4}-\d{2}-\d{2}$", ErrorMessage = "检验日期的格式必须是YYYY-MM-DD")]
        public string CHECKDATE { get; set; }

        [Required(ErrorMessage = "组织内码ID（ORGID）是必填字段")]
        public string ORGID { get; set; }
    }

    public class CheckResultResponse
    {
        public bool Success { get; set; }
        public string Result { get; set; } // 合格、不合格、未检验
        public string Message { get; set; }
    }
}
