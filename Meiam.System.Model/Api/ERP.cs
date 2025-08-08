using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

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

        /// <summary>
        /// 行号
        /// </summary>
        public string SEQ { get; set; }


        [Required(ErrorMessage = "分录内码(ENTRYID)是必填字段")]
        public string ENTRYID { get; set; }

        [Required(ErrorMessage = "收料通知单号(ERP_ARRIVEDID)是必填字段")]
        public string ERP_ARRIVEDID { get; set; }

        [Required(ErrorMessage = "品号(ITEMID)是必填字段")]
        public string ITEMID { get; set; }

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

        // 供应商编码
        public string SUPPCODE { get; set; }

        // 供应商名称
        public string SUPPNAME { get; set; }

        /// <summary>
        /// 金蝶单据类型
        /// </summary>
        public string BUSINESSTYPE { get; set; }
    }

    public class ApiResponse
    {
        public bool Success { get; set; }
        public string Message { get; set; }
        public object Data { get; set; }
    }

    public class ErpApiResponse
    {
        public bool success { get; set; }
        public string message { get; set; }
    }

    /// <summary>
    /// 首检单据
    /// </summary>
    public class WorkOrderSyncRequest
    {
        [Required(ErrorMessage = "工单号（MOID）是必填字段")]
        public string MOID { get; set; }

        [Required(ErrorMessage = "内码（ID）是必填字段")]
        public string ID { get; set; }

        [Required(ErrorMessage = "品号（ITEMID）是必填字段")]
        public string ITEMID { get; set; }
        public string ITEMNAME { get; set; }

        [Required(ErrorMessage = "组织内码ID（ORGID）是必填字段")]
        public string ORGID { get; set; }

        [Required(ErrorMessage = "建单时间（CREATEDATE）是必填字段")]
        [RegularExpression(@"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$",
            ErrorMessage = "建单时间的格式必须是YYYY-MM-DD HH:MM:SS")]
        public string CREATEDATE { get; set; }
    }

    /// <summary>
    /// 收料检验结果回传ERP
    /// </summary>
    public class QmsLotNoticeResultRequest
    {

    }
    /// <summary>
    /// 收料检验结果回传ERP
    /// </summary>
    public class LotNoticeResultRequest
    {
        public string ID { get; set; }

        public string EntryID { get; set; }

        public string BillNo { get; set; }

        /// <summary>
        /// 检验单号
        /// </summary>
        public string InspectBillNo { get; set; }


        public string OrgID { get; set; }

        public string Result { get; set; }

        /// <summary>
        /// 合格数
        /// </summary>
        public int? OKQty { get; set; }

        /// <summary>
        /// 不合格数
        /// </summary>
        public int? NGQty { get; set; }
    }


    /// <summary>
    /// 工单首检检验结果回传MES
    /// </summary>
    public class QmsWorkOrderResultRequest
    {

    }

    /// <summary>
    /// 工单首检检验结果回传MES
    /// </summary>
    public class WorkOrderResultRequest
    {
        public string BillNo { get; set; }

        public string MESFirstInspectID { get; set; }

        public string OrgID { get; set; }

        public string Result { get; set; }

        /// <summary>
        /// 合格数
        /// </summary>
        public int? OKQty { get; set; }

        /// <summary>
        /// 不合格数
        /// </summary>
        public int? NGQty { get; set; }
    }

    /// <summary>
    /// 产品检验结果(入库检)
    /// </summary>
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

    public class CheckResultResponse : ApiResponse
    {
        public string Result { get; set; } // 合格、不合格、未检验
    }

    /// <summary>
    /// ERP物料数据同步
    /// </summary>
    public class MaterialSyncItem
    {
        [Required(ErrorMessage = "品号（ITEMID）是必填字段")]
        public string ITEMID { get; set; }

        public string ITEMNAME { get; set; }

        [Required(ErrorMessage = "组织内码ID（ORGID）是必填字段")]
        public string ORGID { get; set; }
    }

    public class MaterialSyncResponse : ApiResponse
    {
        public int TotalCount { get; set; }
        public int SuccessCount { get; set; }
        public int FailedCount { get; set; }
        public List<MaterialSyncDetail> Details { get; set; } = new();
    }

    public class MaterialSyncDetail
    {
        public string ITEMID { get; set; }
        public string Error { get; set; }
    }

    /// <summary>
    /// ERP客户同步
    /// </summary>
    public class CustomerSyncItem
    {
        public string CUSTOMCODE { get; set; }

        public string CUSTOMNAME { get; set; }

        public string ORGID { get; set; }
    }

    public class CustomerSyncResponse : ApiResponse
    {
        public int TotalCount { get; set; }
        public int SuccessCount { get; set; }
        public int FailedCount { get; set; }
        public List<CustomerSyncDetail> Details { get; set; } = new();
    }

    public class CustomerSyncDetail
    {
        public string CUSTOMCODE { get; set; }
        public string Error { get; set; }
    }

    /// <summary>
    /// ERP供应商同步
    /// </summary>
    public class SuppSyncResponse : ApiResponse
    {
        public int TotalCount { get; set; }
        public int SuccessCount { get; set; }
        public int FailedCount { get; set; }
        public List<SuppSyncDetail> Details { get; set; } = new();
    }

    public class SuppSyncDetail
    {
        public string SUPPID { get; set; }
        public string Error { get; set; }
    }


    public class erp_rc
    {
        public string KEEID { get; set; }
        public string ERP_ARRIVEDID { get; set; }
        public string ITEMID { get; set; }
        public string ITEMNAME { get; set; }
        public decimal LOT_QTY { get; set; }
        public string APPLY_DATE { get; set; }
        public string MODEL_SPEC { get; set; }
        public string LOTNO { get; set; }
        public decimal LENGTH { get; set; }
        public decimal WIDTH { get; set; }
        public decimal INUM { get; set; }
        public string PRO_DATE { get; set; }
        public string QUA_DATE { get; set; }
        public string SUPPNAME { get; set; }
        public string SUPPID { get; set; }
        public string INSPECT_FPICREATEDATE { get; set; }
    }

    public class erp_wr
    {
        public string MOID { get; set; }
        public decimal LOT_QTY { get; set; }
        public decimal REPORT_QTY { get; set; }
        public string ITEMID { get; set; }
        public string ITEMNAME { get; set; }
        public string CREATEDATE { get; set; }
        public string INSPECT02CODE { get; set; }
        public string INSPECT02NAME { get; set; }
        public string INSPECT_FPICREATEDATE { get; set; }
    }

    public class erp_item
    {
        public string ITEMID { get; set; }
        public string ITEMNAME { get; set; }
        public string ITEM_GROUPID { get; set; }
        public string INSPECT_FPICREATEDATE { get; set; }
    }

    public class erp_vend
    {
        public string SUPPNAME { get; set; }
        public string SUPPID { get; set; }
        public string INSPECT_FPICREATEDATE { get; set; }
    }

    public class erp_cust
    {
        public string CUSTOMCODE { get; set; }
        public string CUSTOMNAME { get; set; }
        public string INSPECT_FPICREATEDATE { get; set; }
    }
}
