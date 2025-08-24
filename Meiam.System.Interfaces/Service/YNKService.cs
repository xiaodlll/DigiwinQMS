using DocumentFormat.OpenXml.Spreadsheet;
using Meiam.System.Common;
using Meiam.System.Interfaces.IService;
using Meiam.System.Model;
using Meiam.System.Model.Dto;
using Microsoft.AspNetCore.Components;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using SqlSugar;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
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
                            CASE 
                                WHEN COALESCE(SQM_STATE, OQC_STATE) IN ('OQC_STATE_005', 'OQC_STATE_006', 'OQC_STATE_008') THEN '合格'
                                WHEN COALESCE(SQM_STATE, OQC_STATE) = 'OQC_STATE_007' THEN '不合格'
                                WHEN COALESCE(SQM_STATE, OQC_STATE) = 'OQC_STATE_010' THEN '免检'
                                ELSE '不合格'
                            END AS OQC_STATE,
                            ISNULL(TRY_CAST(FQC_CNT AS INT), 0) AS FQC_CNT,       -- 不可转换时返回0
                            ISNULL(TRY_CAST(FQC_NOT_CNT AS INT), 0) AS FQC_NOT_CNT     -- 不可转换时返回0
                        FROM INSPECT_IQC
                        WHERE (ISSY <> '1' OR ISSY IS NULL) 
                            AND COALESCE(SQM_STATE, OQC_STATE) IN ('OQC_STATE_005', 'OQC_STATE_006', 'OQC_STATE_007', 'OQC_STATE_008', 'OQC_STATE_010')
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

        #region 工具API
        public async Task<ApiResponse> GetAOIInspectInfoByDocCodeAsync(INSPECT_REQCODE input) {
            try {
                var parameters = new SugarParameter[] {
                  new SugarParameter("@DOC_CODE", input.DOC_CODE) };

                var data = await Db.Ado.SqlQueryAsync<INSPECT_INFO_BYCODE>(
                    "select TOP 1 ITEMID,ITEMNAME,LOTNO,LOT_QTY from INSPECT_VIEW where INSPECT_CODE = @DOC_CODE",
                    parameters
                );

                return new ApiResponse {
                    Success = true,
                    Message = "数据获取成功",
                    Data = JsonConvert.SerializeObject(data.FirstOrDefault())
                };
            }
            catch (Exception ex) {
                return new ApiResponse {
                    Success = false,
                    Message = $"数据获取失败：{ex.Message}"
                };
            }
        }

        public async Task<ApiResponse> GetAOIProgressDataByDocCodeAsync(INSPECT_REQCODE input) {
            try {
                var parameters = new SugarParameter[] {
                  new SugarParameter("@DOC_CODE", input.DOC_CODE) };

                List<INSPECT_PROGRESS_BYCODE> data = await Db.Ado.SqlQueryAsync<INSPECT_PROGRESS_BYCODE>(
                    "select INSPECT_PROGRESSNAME, INSPECT_PROGRESSID from INSPECT_PROGRESS where DOC_CODE = @DOC_CODE and INSPECT_DEV='INSPECT_DEV_010'",
                    parameters
                );

                return new ApiResponse {
                    Success = true,
                    Message = "数据获取成功",
                    Data = JsonConvert.SerializeObject(data)
                };
            }
            catch (Exception ex) {
                return new ApiResponse {
                    Success = false,
                    Message = $"数据获取失败：{ex.Message}"
                };
            }
        }

        public async Task<ApiResponse> ProcessUploadAOIDataAsync(List<InspectAoi> input) {
            try {
                foreach (var item in input) {
                    // 1. 检查主表是否存在相同数据
                    var checkSql = @"SELECT COUNT(1) FROM INSPECT_AOI 
                                WHERE DOC_CODE = @DOC_CODE 
                                  AND InspectionDate = @InspectionDate 
                                  AND BeginTime = @BeginTime
                                  AND ComponentName = @ComponentName
                                  AND MainSN = @MainSN
                                  AND PanelSN = @PanelSN
                                  AND PanelID = @PanelID"
                    ;

                    var checkParams = new SugarParameter[]
                    {
                    new SugarParameter("@DOC_CODE", item.DOC_CODE),
                    new SugarParameter("@InspectionDate", item.InspectionDate),
                    new SugarParameter("@BeginTime", item.BeginTime),
                    new SugarParameter("@ComponentName", item.ComponentName),
                    new SugarParameter("@MainSN", item.MainSN),
                    new SugarParameter("@PanelSN", item.PanelSN),
                    new SugarParameter("@PanelID", item.PanelID)
                    };

                    // 执行查询，判断是否存在
                    var exists = await Db.Ado.GetIntAsync(checkSql, checkParams) > 0;

                    // 2数据处理
                    if (!exists) {
                        // 构建插入SQL语句
                        var mainSql = @"INSERT INTO INSPECT_AOI 
                            (DOC_CODE, INSPECT_PROGRESSID, INSPECT_AOIID, INSPECT_AOICREATEDATE, 
                             INSPECT_AOICREATEUSER, TENID, MainSN, PanelSN, PanelID, ModelName, 
                             Side, MachineName, CustomerName, Operator, Programer, InspectionDate, 
                             BeginTime, EndTime, CycleTimeSec, InspectionBatch, ReportResult, 
                             ConfirmedResult, TotalComponent, ReportFailComponent, ComfirmedFailComponent, 
                             ComponentName, LibraryModel, PN, Package, Angle, NGReportResult, 
                             ReportResultCode, NGConfirmedResult, ConfirmedResultCode)
                            VALUES 
                            (@DOC_CODE, @INSPECT_PROGRESSID, @INSPECT_AOIID, @INSPECT_AOICREATEDATE, 
                             @INSPECT_AOICREATEUSER, @TENID, @MainSN, @PanelSN, @PanelID, @ModelName, 
                             @Side, @MachineName, @CustomerName, @Operator, @Programer, @InspectionDate, 
                             @BeginTime, @EndTime, @CycleTimeSec, @InspectionBatch, @ReportResult, 
                             @ConfirmedResult, @TotalComponent, @ReportFailComponent, @ComfirmedFailComponent, 
                             @ComponentName, @LibraryModel, @PN, @Package, @Angle, @NGReportResult, 
                             @ReportResultCode, @NGConfirmedResult, @ConfirmedResultCode)";

                        // 构建参数数组，与实体属性一一对应
                        var mainParams = new SugarParameter[]
                        {
                    new SugarParameter("@DOC_CODE", item.DOC_CODE),
                    new SugarParameter("@INSPECT_PROGRESSID", item.INSPECT_PROGRESSID),
                    new SugarParameter("@INSPECT_AOIID", item.INSPECT_AOIID),
                    new SugarParameter("@INSPECT_AOICREATEDATE", item.INSPECT_AOICREATEDATE),
                    new SugarParameter("@INSPECT_AOICREATEUSER", item.INSPECT_AOICREATEUSER),
                    new SugarParameter("@TENID", item.TENID),
                    new SugarParameter("@MainSN", item.MainSN),
                    new SugarParameter("@PanelSN", item.PanelSN),
                    new SugarParameter("@PanelID", item.PanelID),
                    new SugarParameter("@ModelName", item.ModelName),
                    new SugarParameter("@Side", item.Side),
                    new SugarParameter("@MachineName", item.MachineName),
                    new SugarParameter("@CustomerName", item.CustomerName),
                    new SugarParameter("@Operator", item.Operator),
                    new SugarParameter("@Programer", item.Programer),
                    new SugarParameter("@InspectionDate", item.InspectionDate),
                    new SugarParameter("@BeginTime", item.BeginTime),
                    new SugarParameter("@EndTime", item.EndTime),
                    new SugarParameter("@CycleTimeSec", item.CycleTimeSec),
                    new SugarParameter("@InspectionBatch", item.InspectionBatch),
                    new SugarParameter("@ReportResult", item.ReportResult),
                    new SugarParameter("@ConfirmedResult", item.ConfirmedResult),
                    new SugarParameter("@TotalComponent", item.TotalComponent),
                    new SugarParameter("@ReportFailComponent", item.ReportFailComponent),
                    new SugarParameter("@ComfirmedFailComponent", item.ComfirmedFailComponent),
                    new SugarParameter("@ComponentName", item.ComponentName),
                    new SugarParameter("@LibraryModel", item.LibraryModel),
                    new SugarParameter("@PN", item.PN),
                    new SugarParameter("@Package", item.Package),
                    new SugarParameter("@Angle", item.Angle),
                    new SugarParameter("@NGReportResult", item.NGReportResult),
                    new SugarParameter("@ReportResultCode", item.ReportResultCode),
                    new SugarParameter("@NGConfirmedResult", item.NGConfirmedResult),
                    new SugarParameter("@ConfirmedResultCode", item.ConfirmedResultCode)
                        };

                        // 执行插入操作
                        await Db.Ado.ExecuteCommandAsync(mainSql, mainParams);
                    }
                }

                return new ApiResponse {
                    Success = true,
                    Message = "Aoi数据保存成功"
                };
            }
            catch (Exception ex) {
                return new ApiResponse {
                    Success = false,
                    Message = $"Aoi数据保存失败：{ex.Message}"
                };
            }
        }

        public async Task<ApiResponse> ProcessUploadAOIImageDataAsync(List<InspectImageAoi> input) {
            try {
                if (input.Count == 0) {
                    return new ApiResponse {
                        Success = false,
                        Message = $"传入数据为空!"
                    };
                }
                string baseDirPath = Path.Combine(AppSettings.Configuration["AppSettings:FileServerPath"], @"AOI");
                foreach (var item in input) {
                    // 1. 验证必要数据
                    if (string.IsNullOrEmpty(item.DOC_CODE)) {
                        throw new ArgumentException("DOC_CODE不能为空，无法存储图片");
                    }
                    if (string.IsNullOrEmpty(item.ImageName)) {
                        throw new ArgumentException("ImageName不能为空，无法确定文件名");
                    }
                    if (string.IsNullOrEmpty(item.ImageData)) {
                        throw new ArgumentException($"DOC_CODE: {item.DOC_CODE} 的图片数据为空，无法存储");
                    }

                    // 2. 创建DOC_CODE对应的文件夹
                    string docCodeDirPath = Path.Combine(baseDirPath, item.DOC_CODE);
                    if (!Directory.Exists(docCodeDirPath)) {
                        Directory.CreateDirectory(docCodeDirPath);
                    }

                    // 3. 处理文件名（存在时添加_1、_2等后缀）
                    string fileExtension = Path.GetExtension(item.ImageName);
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(item.ImageName);
                    string targetFilePath = Path.Combine(docCodeDirPath, item.ImageName);

                    // 检查文件是否存在，存在则生成新文件名
                    if (File.Exists(targetFilePath)) {
                        int counter = 1;
                        do {
                            string newFileName = $"{fileNameWithoutExt}_{counter}{fileExtension}";
                            targetFilePath = Path.Combine(docCodeDirPath, newFileName);
                            counter++;
                        } while (File.Exists(targetFilePath));
                    }

                    // 4. 将Base64字符串转换为图片并保存
                    byte[] imageBytes = Convert.FromBase64String(item.ImageData);
                    await File.WriteAllBytesAsync(targetFilePath, imageBytes);

                    // 5. 更新实体中的ImageName为实际保存的文件名（如果有重命名）
                    item.ImageName = Path.GetFileName(targetFilePath);
                }

                return new ApiResponse {
                    Success = true,
                    Message = "AOI图片保存成功",
                    Data = JsonConvert.SerializeObject(input)
                };
            }
            catch (Exception ex) {
                return new ApiResponse {
                    Success = false,
                    Message = $"AOI图片数据保存失败：{ex.Message}"
                };
            }
        }

        #endregion
    }
}
