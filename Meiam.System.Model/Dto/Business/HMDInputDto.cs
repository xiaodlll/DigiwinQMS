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

    /// <summary>
    /// 恒铭达拉力机主表
    /// </summary>
    public class InspectDev1Entity {
        [Display(Name = "记录唯一ID")]
        public string INSPECT_DEV1ID { get; set; }

        [Display(Name = "检验单号")]
        public string INSPECT_CODE { get; set; }

        [Display(Name = "检验项目")]
        public string INSPECT_PROGRESSID { get; set; }

        [Display(Name = "是否关联")]
        public string ISBUILD { get; set; }

        [Display(Name = "检验规格")]
        public string INSPECT_SPEC { get; set; }

        [Display(Name = "物料编码")]
        public string ITEMID { get; set; }

        [Display(Name = "物料名称")]
        public string ITEMNAME { get; set; }

        [Display(Name = "检验项目名称")]
        public string INSPECTTYPE1 { get; set; }

        [Display(Name = "批号")]
        public string LOTID { get; set; }

        [Display(Name = "批次数量")]
        public string LOT_QTY { get; set; }

        [Display(Name = "实际样本数")]
        public string SAMPLE_CNT { get; set; }

        [Display(Name = "批次标识")]
        public string BATCHID { get; set; }

        [Display(Name = "剥离区间1起点")]
        public string DEFORMATION_START { get; set; }

        [Display(Name = "剥离区间1终点")]
        public string DEFORMATION_END { get; set; }

        [Display(Name = "剥离区间2起点")]
        public string DEFORMATION_START2 { get; set; }

        [Display(Name = "剥离区间2终点")]
        public string DEFORMATION_END2 { get; set; }

        [Display(Name = "剥离区间3起点")]
        public string DEFORMATION_START3 { get; set; }

        [Display(Name = "剥离区间3终点")]
        public string DEFORMATION_END3 { get; set; }

        [Display(Name = "剥离区间4起点")]
        public string DEFORMATION_START4 { get; set; }

        [Display(Name = "剥离区间4终点")]
        public string DEFORMATION_END4 { get; set; }

        [Display(Name = "剥离区间5起点")]
        public string DEFORMATION_START5 { get; set; }

        [Display(Name = "剥离区间5终点")]
        public string DEFORMATION_END5 { get; set; }

        [Display(Name = "检验人员")]
        public string PEOPLE02 { get; set; }

        [Display(Name = "检验人员")]
        public string APPEOPLE02 { get; set; }

        [Display(Name = "租户ID")]
        public string TENID { get; set; }

        [Display(Name = "创建用户")]
        public string INSPECT_DEV1CREATEUSER { get; set; }

        [Display(Name = "创建时间")]
        public string INSPECT_DEV1CREATEDATE { get; set; }

        public List<InspectTensileEntity> Details { get; set; } = new List<InspectTensileEntity>();
    }

    /// <summary>
    /// 恒铭达拉力机明细表
    /// </summary>
    public class InspectTensileEntity {

        [Display(Name = "记录唯一ID")]
        public string INSPECT_TENSILEID { get; set; }

        [Display(Name = "产品类型")]
        public string ITEMNAME { get; set; }

        [Display(Name = "测试批号")]
        public string TESTLOT { get; set; }

        [Display(Name = "测试类型")]
        public string TESTTYPE { get; set; }

        [Display(Name = "测试角度")]
        public string INSPECTTYPE1 { get; set; }

        [Display(Name = "测量时间")]
        public string INSPECT_DATE { get; set; }

        [Display(Name = "测试人员")]
        public string PEOPLE02 { get; set; }

        [Display(Name = "审批人员")]
        public string APPEOPLE02 { get; set; }

        [Display(Name = "X轴值")]
        public string X_AXIS { get; set; }

        [Display(Name = "Y轴值")]
        public string Y_AXIS { get; set; }

        [Display(Name = "样本编码")]
        public string SAMPLEID { get; set; }

        [Display(Name = "同一次测量的标识")]
        public string BATCHID { get; set; }

        [Display(Name = "断裂伸长率")]
        public string EAB { get; set; }

        [Display(Name = "拉伸强度")]
        public string MPA { get; set; }

        [Display(Name = "面积")]
        public string AREA { get; set; }

        [Display(Name = "厚度")]
        public string THICKNESS { get; set; }

        [Display(Name = "宽度")]
        public string WIDTH { get; set; }

        [Display(Name = "剥离区间1起点")]
        public string DEFORMATION_START { get; set; }

        [Display(Name = "剥离区间1终点")]
        public string DEFORMATION_END { get; set; }

        [Display(Name = "剥离区间2起点")]
        public string DEFORMATION_START2 { get; set; }

        [Display(Name = "剥离区间2终点")]
        public string DEFORMATION_END2 { get; set; }

        [Display(Name = "剥离区间3起点")]
        public string DEFORMATION_START3 { get; set; }

        [Display(Name = "剥离区间3终点")]
        public string DEFORMATION_END3 { get; set; }

        [Display(Name = "剥离区间4起点")]
        public string DEFORMATION_START4 { get; set; }

        [Display(Name = "剥离区间4终点")]
        public string DEFORMATION_END4 { get; set; }

        [Display(Name = "剥离区间5起点")]
        public string DEFORMATION_START5 { get; set; }

        [Display(Name = "剥离区间5终点")]
        public string DEFORMATION_END5 { get; set; }

        [Display(Name = "主表ID")]
        public string INSPECT_DEV1ID { get; set; }

        [Display(Name = "租户ID")]
        public string TENID { get; set; }

        [Display(Name = "创建用户")]
        public string INSPECT_TENSILECREATEUSER { get; set; }

        [Display(Name = "创建时间")]
        public string INSPECT_TENSILECREATEDATE { get; set; }
    }

    public class INSPECT_PROGRESSDto {
        /// <summary>
        /// 主键ID
        /// </summary>
        [Display(Name = "主键ID")]
        public string INSPECT_PROGRESSID { get; set; }

        /// <summary>
        /// 检验单号
        /// </summary>
        [Display(Name = "检验单号")]
        public string DOC_CODE { get; set; }

        /// <summary>
        /// 品号
        /// </summary>
        [Display(Name = "品号")]
        public string ITEMID { get; set; }

        /// <summary>
        /// 版本
        /// </summary>
        [Display(Name = "版本")]
        public string VER { get; set; }

        /// <summary>
        /// 顺序号
        /// </summary>
        [Display(Name = "顺序号")]
        public string OID { get; set; }

        /// <summary>
        /// 属性，默认值：COC_ATTR_001
        /// </summary>
        [Display(Name = "属性")]
        public string COC_ATTR { get; set; } = "COC_ATTR_001";

        /// <summary>
        /// 检验项目
        /// </summary>
        [Display(Name = "检验项目")]
        public string INSPECT_PROGRESSNAME { get; set; }

        /// <summary>
        /// 检验仪器，默认值：INSPECT_DEV_001
        /// </summary>
        [Display(Name = "检验仪器")]
        public string INSPECT_DEV { get; set; } = "INSPECT_DEV_001";

        /// <summary>
        /// 分析方法，默认值：COUNTTYPE_002
        /// </summary>
        [Display(Name = "分析方法")]
        public string COUNTTYPE { get; set; } = "COUNTTYPE_002";

        /// <summary>
        /// 检验标准
        /// </summary>
        [Display(Name = "检验标准")]
        public string INSPECT_PLANID { get; set; }

        /// <summary>
        /// 应检数量
        /// </summary>
        [Display(Name = "应检数量")]
        public string INSPECT_CNT { get; set; }

        /// <summary>
        /// 标准值
        /// </summary>
        [Display(Name = "标准值")]
        public string STD_VALUE { get; set; }

        /// <summary>
        /// 上公差
        /// </summary>
        [Display(Name = "上公差")]
        public string MAX_VALUE { get; set; }

        /// <summary>
        /// 下公差
        /// </summary>
        [Display(Name = "下公差")]
        public string MIN_VALUE { get; set; }

        /// <summary>
        /// 上限值
        /// </summary>
        [Display(Name = "上限值")]
        public string UP_VALUE { get; set; }

        /// <summary>
        /// 下限值
        /// </summary>
        [Display(Name = "下限值")]
        public string DOWN_VALUE { get; set; }

        /// <summary>
        /// 样本1
        /// </summary>
        [Display(Name = "样本1")]
        public string A1 { get; set; }

        /// <summary>
        /// 样本2
        /// </summary>
        [Display(Name = "样本2")]
        public string A2 { get; set; }

        /// <summary>
        /// 样本3
        /// </summary>
        [Display(Name = "样本3")]
        public string A3 { get; set; }

        /// <summary>
        /// 样本4
        /// </summary>
        [Display(Name = "样本4")]
        public string A4 { get; set; }

        /// <summary>
        /// 样本5
        /// </summary>
        [Display(Name = "样本5")]
        public string A5 { get; set; }

        /// <summary>
        /// 样本6
        /// </summary>
        [Display(Name = "样本6")]
        public string A6 { get; set; }

        /// <summary>
        /// 样本7
        /// </summary>
        [Display(Name = "样本7")]
        public string A7 { get; set; }

        /// <summary>
        /// 样本8
        /// </summary>
        [Display(Name = "样本8")]
        public string A8 { get; set; }

        /// <summary>
        /// 样本9
        /// </summary>
        [Display(Name = "样本9")]
        public string A9 { get; set; }

        /// <summary>
        /// 样本10
        /// </summary>
        [Display(Name = "样本10")]
        public string A10 { get; set; }

        /// <summary>
        /// 样本11
        /// </summary>
        [Display(Name = "样本11")]
        public string A11 { get; set; }

        /// <summary>
        /// 样本12
        /// </summary>
        [Display(Name = "样本12")]
        public string A12 { get; set; }

        /// <summary>
        /// 样本13
        /// </summary>
        [Display(Name = "样本13")]
        public string A13 { get; set; }

        /// <summary>
        /// 样本14
        /// </summary>
        [Display(Name = "样本14")]
        public string A14 { get; set; }

        /// <summary>
        /// 样本15
        /// </summary>
        [Display(Name = "样本15")]
        public string A15 { get; set; }

        /// <summary>
        /// 样本16
        /// </summary>
        [Display(Name = "样本16")]
        public string A16 { get; set; }

        /// <summary>
        /// 样本17
        /// </summary>
        [Display(Name = "样本17")]
        public string A17 { get; set; }

        /// <summary>
        /// 样本18
        /// </summary>
        [Display(Name = "样本18")]
        public string A18 { get; set; }

        /// <summary>
        /// 样本19
        /// </summary>
        [Display(Name = "样本19")]
        public string A19 { get; set; }

        /// <summary>
        /// 样本20
        /// </summary>
        [Display(Name = "样本20")]
        public string A20 { get; set; }

        /// <summary>
        /// 样本21
        /// </summary>
        [Display(Name = "样本21")]
        public string A21 { get; set; }

        /// <summary>
        /// 样本22
        /// </summary>
        [Display(Name = "样本22")]
        public string A22 { get; set; }

        /// <summary>
        /// 样本23
        /// </summary>
        [Display(Name = "样本23")]
        public string A23 { get; set; }

        /// <summary>
        /// 样本24
        /// </summary>
        [Display(Name = "样本24")]
        public string A24 { get; set; }

        /// <summary>
        /// 样本25
        /// </summary>
        [Display(Name = "样本25")]
        public string A25 { get; set; }

        /// <summary>
        /// 样本26
        /// </summary>
        [Display(Name = "样本26")]
        public string A26 { get; set; }

        /// <summary>
        /// 样本27
        /// </summary>
        [Display(Name = "样本27")]
        public string A27 { get; set; }

        /// <summary>
        /// 样本28
        /// </summary>
        [Display(Name = "样本28")]
        public string A28 { get; set; }

        /// <summary>
        /// 样本29
        /// </summary>
        [Display(Name = "样本29")]
        public string A29 { get; set; }

        /// <summary>
        /// 样本30
        /// </summary>
        [Display(Name = "样本30")]
        public string A30 { get; set; }

        /// <summary>
        /// 样本31
        /// </summary>
        [Display(Name = "样本31")]
        public string A31 { get; set; }

        /// <summary>
        /// 样本32
        /// </summary>
        [Display(Name = "样本32")]
        public string A32 { get; set; }

        /// <summary>
        /// 样本33
        /// </summary>
        [Display(Name = "样本33")]
        public string A33 { get; set; }

        /// <summary>
        /// 样本34
        /// </summary>
        [Display(Name = "样本34")]
        public string A34 { get; set; }

        /// <summary>
        /// 样本35
        /// </summary>
        [Display(Name = "样本35")]
        public string A35 { get; set; }

        /// <summary>
        /// 样本36
        /// </summary>
        [Display(Name = "样本36")]
        public string A36 { get; set; }

        /// <summary>
        /// 样本37
        /// </summary>
        [Display(Name = "样本37")]
        public string A37 { get; set; }

        /// <summary>
        /// 样本38
        /// </summary>
        [Display(Name = "样本38")]
        public string A38 { get; set; }

        /// <summary>
        /// 样本39
        /// </summary>
        [Display(Name = "样本39")]
        public string A39 { get; set; }

        /// <summary>
        /// 样本40
        /// </summary>
        [Display(Name = "样本40")]
        public string A40 { get; set; }

        /// <summary>
        /// 样本41
        /// </summary>
        [Display(Name = "样本41")]
        public string A41 { get; set; }

        /// <summary>
        /// 样本42
        /// </summary>
        [Display(Name = "样本42")]
        public string A42 { get; set; }

        /// <summary>
        /// 样本43
        /// </summary>
        [Display(Name = "样本43")]
        public string A43 { get; set; }

        /// <summary>
        /// 样本44
        /// </summary>
        [Display(Name = "样本44")]
        public string A44 { get; set; }

        /// <summary>
        /// 样本45
        /// </summary>
        [Display(Name = "样本45")]
        public string A45 { get; set; }

        /// <summary>
        /// 样本46
        /// </summary>
        [Display(Name = "样本46")]
        public string A46 { get; set; }

        /// <summary>
        /// 样本47
        /// </summary>
        [Display(Name = "样本47")]
        public string A47 { get; set; }

        /// <summary>
        /// 样本48
        /// </summary>
        [Display(Name = "样本48")]
        public string A48 { get; set; }

        /// <summary>
        /// 样本49
        /// </summary>
        [Display(Name = "样本49")]
        public string A49 { get; set; }

        /// <summary>
        /// 样本50
        /// </summary>
        [Display(Name = "样本50")]
        public string A50 { get; set; }

        /// <summary>
        /// 样本51
        /// </summary>
        [Display(Name = "样本51")]
        public string A51 { get; set; }

        /// <summary>
        /// 样本52
        /// </summary>
        [Display(Name = "样本52")]
        public string A52 { get; set; }

        /// <summary>
        /// 样本53
        /// </summary>
        [Display(Name = "样本53")]
        public string A53 { get; set; }

        /// <summary>
        /// 样本54
        /// </summary>
        [Display(Name = "样本54")]
        public string A54 { get; set; }

        /// <summary>
        /// 样本55
        /// </summary>
        [Display(Name = "样本55")]
        public string A55 { get; set; }

        /// <summary>
        /// 样本56
        /// </summary>
        [Display(Name = "样本56")]
        public string A56 { get; set; }

        /// <summary>
        /// 样本57
        /// </summary>
        [Display(Name = "样本57")]
        public string A57 { get; set; }

        /// <summary>
        /// 样本58
        /// </summary>
        [Display(Name = "样本58")]
        public string A58 { get; set; }

        /// <summary>
        /// 样本59
        /// </summary>
        [Display(Name = "样本59")]
        public string A59 { get; set; }

        /// <summary>
        /// 样本60
        /// </summary>
        [Display(Name = "样本60")]
        public string A60 { get; set; }

        /// <summary>
        /// 样本61
        /// </summary>
        [Display(Name = "样本61")]
        public string A61 { get; set; }

        /// <summary>
        /// 样本62
        /// </summary>
        [Display(Name = "样本62")]
        public string A62 { get; set; }

        /// <summary>
        /// 样本63
        /// </summary>
        [Display(Name = "样本63")]
        public string A63 { get; set; }

        /// <summary>
        /// 样本64
        /// </summary>
        [Display(Name = "样本64")]
        public string A64 { get; set; }

        /// <summary>
        /// 用户名
        /// </summary>
        [Display(Name = "用户名")]
        public string INSPECT_PROGRESSCREATEUSER { get; set; }

        /// <summary>
        /// 当前日期
        /// </summary>
        [Display(Name = "当前日期")]
        public string INSPECT_PROGRESSDATE { get; set; }

        /// <summary>
        /// TENID，默认值：001
        /// </summary>
        [Display(Name = "TENID")]
        public string TENID { get; set; } = "001";
    }

    public class INSPECT_REQCODE {
        public string DOC_CODE { get; set; }
    }

    public class INSPECT_CONDITION {
        public string DOC_CODE { get; set; }
        public string ITEMID { get; set; }
        public string ITEMNAME { get; set; }
        public string LOTNO { get; set; }
        public string LOT_QTY { get; set; }
    }

    public class INSPECT_INFO_BYCODE {
        public string INSPECT_CODE { get; set; }
        public string ITEMID { get; set; }
        public string ITEMNAME { get; set; }
        public string LOTNO { get; set; }
        public string LOT_QTY { get; set; }
    }

    public class INSPECT_PROGRESS_BYCODE {
        public string INSPECT_PROGRESSID { get; set; }
        public string INSPECT_PROGRESSNAME { get; set; }
    }

    public class INSPECT_SYSM002_REQBYID {
        public string SYSM001ID { get; set; }
    }

    public class INSPECT_SYSM002_BYID {
        public string SYSM002ID { get; set; }
        public string SYSM002NAME { get; set; }
    }
}
