using Meiam.System.Interfaces.IService;
using Meiam.System.Model;
using Microsoft.Extensions.Logging;
using SqlSugar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Meiam.System.Interfaces.Service
{
    public class YNKService : BaseService<INSPECT_TENSILE_D>, IYNKService
    {
        public YNKService(IUnitOfWork unitOfWork) : base(unitOfWork)
        {
        }

        private readonly ILogger<YNKService> _logger;
        private readonly string _connectionString;

        public YNKService(IUnitOfWork unitOfWork, ILogger<YNKService> logger) : base(unitOfWork)
        {
            _logger = logger;
        }

        #region ERP收料通知单
        public async Task<ApiResponse> ProcessLotNoticeAsync(List<LotNoticeRequest> requests)
        {
            _logger.LogInformation("开始处理收料通知单");
            try
            {
                foreach (var request in requests)
                {
                    // 验证数据
                    if (request.LOT_QTY <= 0)
                    {
                        _logger.LogWarning("到货数量无效: {LotQty}", request.LOT_QTY);
                        throw new ArgumentException("到货数量必须大于0");
                    }

                    //判断重复
                    bool isExist = Db.Ado.GetInt($@"SELECT count(*) FROM INSPECT_IQC WHERE KEEID = '{request.ID}'") > 0;
                    if (isExist)
                    {
                        _logger.LogWarning($"收料通知单已存在: {request.ID}");
                        continue;
                    }

                    // 生成检验单号
                    var inspectionId = GenerateInspectionId();
                    _logger.LogInformation("生成检验单号: {InspectionId}", inspectionId);

                    // 保存到数据库
                    _logger.LogDebug("正在保存收料通知单到数据库...");
                    try
                    {
                        SaveToDatabase(request, inspectionId);
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError("保存收料通知单到数据库异常:" + ex.ToString());
                        throw;
                    }
                    await Task.Delay(10); // 模拟异步操作
                }

                _logger.LogInformation("收料通知单处理成功");

                return new ApiResponse
                {
                    Success = true,
                    Message = "收料通知单接收成功",
                };
            }
            catch (Exception ex)
            {
                return new ApiResponse
                {
                    Success = false,
                    Message = $"收料通知单接收失败，原因：{ex.Message}"
                };
            }
        }

        private string GenerateInspectionId()
        {
            string INSPECT_CODE = "";//检验单号

            const string sql = @"
                DECLARE @INSPECT_CODE  	  NVARCHAR(200) 

                --获得IQC检验单号
                SELECT TOP 1 @INSPECT_CODE=CAST(CAST(dbo.getNumericValue(INSPECT_IQCCODE) AS DECIMAL)+1 AS CHAR)  FROM  INSPECT_IQC
                WHERE  TENID='001' AND ISNULL(REPLACE(INSPECT_IQCCODE,'IQC_',''),'') like REPLACE(CONVERT(VARCHAR(10),GETDATE(),120),'-','')+'%' 
                ORDER BY INSPECT_IQCCODE DESC

                IF(ISNULL(@INSPECT_CODE,'')='')
                   SET @INSPECT_CODE ='IQC_'+REPLACE(CONVERT(VARCHAR(10),GETDATE(),120),'-','')+'001'
                ELSE 
                   SET @INSPECT_CODE ='IQC_'+@INSPECT_CODE

                SELECT @INSPECT_CODE AS INSPECT_CODE
                ";

            // 执行 SQL 命令
            var dataTable = Db.Ado.GetDataTable(sql);
            if (dataTable.Rows.Count > 0)
            {
                INSPECT_CODE = dataTable.Rows[0]["INSPECT_CODE"].ToString().Trim();
            }
            return INSPECT_CODE;
        }

        public void SaveToDatabase(LotNoticeRequest request, string inspectionId)
        {
            // 保存数据
            SaveMainInspection(request, inspectionId);
        }

        private void SaveMainInspection(LotNoticeRequest request, string inspectionId)
        {
            //更新供应商SUPP
            string SuppID = Db.Ado.GetScalar($@"SELECT TOP 1 SUPPID FROM SUPP WHERE SUPPNAME = '{request.SUPPNAME}'")?.ToString().Trim();

            if (string.IsNullOrEmpty(SuppID))
            {
                SuppID = Db.Ado.GetScalar($@"select TOP 1 cast(cast(dbo.getNumericValue(SUPPID) AS DECIMAL)+1 as char) from SUPP order by SUPPID desc")?.ToString().Trim();
                if (string.IsNullOrEmpty(SuppID))
                {
                    SuppID = "1001";
                }
                Db.Ado.ExecuteCommand($@"INSERT INTO SUPP (
                                            TENID, SUPPID, SUPP0A17, SUPPCREATEUSER, SUPPCREATEDATE,
                                            SUPPMODIFYDATE, SUPPMODIFYUSER, SUPPCODE, SUPPNAME)
                                        VALUES (
                                            '001', '{SuppID}', '001', 'system', '{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}',
                                            '{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}', 'system', '{SuppID}', '{request.SUPPNAME}')");
            }

            //更新物料ITEM
            string ItemID = Db.Ado.GetScalar($@"SELECT TOP 1 ITEMID FROM ITEM WHERE ITEMNAME = '{request.ITEMNAME}'")?.ToString().Trim();

            if (string.IsNullOrEmpty(ItemID))
            {
                ItemID = Db.Ado.GetScalar($@"select TOP 1 cast(cast(dbo.getNumericValue(ITEMID) AS DECIMAL)+1 as char) from ITEM order by ITEMID desc")?.ToString().Trim();
                if (string.IsNullOrEmpty(ItemID))
                {
                    ItemID = "1001";
                }
                Db.Ado.ExecuteCommand($@"INSERT INTO ITEM (
                                            TENID, ITEMID, ITEM0A17, ITEMCREATEUSER, ITEMCREATEDATE,
                                            ITEMMODIFYDATE, ITEMMODIFYUSER, ITEMCODE, ITEMNAME)
                                        VALUES (
                                            '001', '{ItemID}', '001', 'system', '{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}',
                                            '{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}', 'system', '{ItemID}', '{request.ITEMNAME}')");
            }

            //更新INSPECT_IQC
            string sql = @"
                INSERT INTO INSPECT_IQC (
                    TENID, INSPECT_IQCID, INSPECT_IQCCREATEUSER, 
                    INSPECT_IQCCREATEDATE, ITEMNAME, ERP_ARRIVEDID, 
                    LOT_QTY, INSPECT_IQCCODE, ITEMID, LOTNO, 
                    APPLY_DATE, ITEM_SPECIFICATION, UNIT
                    SUPPID, SUPPNAME, SUPPLOTNO, KEEID
                ) VALUES (
                    @TenId, @InspectIqcId, @InspectIqcCreateUser, 
                    getdate(), @ItemName, @ErpArrivedId,
                    @LotQty, @InspectIqcCode, @ItemId, @LotNo, 
                    @ApplyDate, @ItemSpecification, @Unit
                    @SuppID, @SuppName, @SuppLotNo, @KeeId
                )";

            // 定义参数
            var parameters = new SugarParameter[]
            {
                new SugarParameter("@TenId", "001"),
                new SugarParameter("@InspectIqcId", inspectionId),
                new SugarParameter("@InspectIqcCreateUser", "system"),
                new SugarParameter("@ItemName", request.ITEMNAME),
                new SugarParameter("@ErpArrivedId", request.ERP_ARRIVEDID),
                new SugarParameter("@LotQty", request.LOT_QTY),
                new SugarParameter("@InspectIqcCode", inspectionId),
                new SugarParameter("@ItemId", ItemID),
                new SugarParameter("@LotNo", (request.LOTNO==null?"":request.LOTNO.ToString())),
                new SugarParameter("@ApplyDate", request.APPLY_DATE),
                new SugarParameter("@ItemSpecification", request.MODEL_SPEC),
                new SugarParameter("@Unit", request.UNIT),
                new SugarParameter("@SuppID", SuppID),
                new SugarParameter("@SuppName", request.SUPPNAME),
                new SugarParameter("@SuppLotNo", request.SUPPLOTNO),
                new SugarParameter("@KeeId", request.ID)
            };

            // 执行 SQL 命令
            Db.Ado.ExecuteCommand(sql, parameters);
        }
        #endregion


        #region 回写ERP MES方法
        public List<LotNoticeResultRequest> GetQmsLotNoticeResultRequest()
        {
            var sql = @"SELECT TOP 100 
                            ERP_ARRIVEDID AS ERP_ARRIVEDID, 
                            INSPECT_IQCCODE AS INSPECT_IQCCODE,
                            ITEMID AS ITEMID,
                            ITEMNAME AS ITEMNAME,
                            LOTNO AS LOTNO,
                            KEEID AS KEEID,
                            CASE WHEN OQC_STATE IN ('OQC_STATE_005', 'OQC_STATE_006', 'OQC_STATE_008') THEN 1 ELSE 0 END AS OQC_STATE,
                            ISNULL(TRY_CAST(FQC_CNT AS INT), 0) AS FQC_CNT,       -- 不可转换时返回0
                            ISNULL(TRY_CAST(FQC_NOT_CNT AS INT), 0) AS FQC_NOT_CNT     -- 不可转换时返回0
                        FROM INSPECT_IQC
                        WHERE (ISSY <> '1' OR ISSY IS NULL) 
                            AND OQC_STATE IN ('OQC_STATE_005', 'OQC_STATE_006', 'OQC_STATE_007', 'OQC_STATE_008')
                        ORDER BY INSPECT_IQCCREATEDATE DESC;";

            var list = Db.Ado.SqlQuery<LotNoticeResultRequest>(sql);
            return list;
        }

        public void CallBackQmsLotNoticeResult(LotNoticeResultRequest request)
        {
            var sql = string.Format(@"update INSPECT_IQC set ISSY='1' where KEEID='{0}' ", request.ID);
            Db.Ado.ExecuteCommand(sql);
        }
        #endregion
    }
}
