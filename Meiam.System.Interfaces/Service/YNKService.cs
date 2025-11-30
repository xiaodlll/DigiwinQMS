using Aspose.Pdf.Operators;
using DocumentFormat.OpenXml.Spreadsheet;
using Meiam.System.Common;
using Meiam.System.Interfaces.IService;
using Meiam.System.Model;
using Meiam.System.Model.Dto;
using Microsoft.AspNetCore.Components;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SqlSugar;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.PerformanceData;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection.Emit;
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
        private readonly ERPApiConfigYNK _config;
        private readonly HttpClient _httpClient;

        public YNKService(IUnitOfWork unitOfWork, HttpClient httpClient, IConfiguration configuration, ILogger<YNKService> logger) : base(unitOfWork)
        {
            _logger = logger;

            _httpClient = httpClient;
            _config = new ERPApiConfigYNK
            {
                BaseUrl = configuration["ERP:BaseUrl"] ?? "http://kingdeeapp/K3Cloud/",
                LoginUrl = configuration["ERP:LoginUrl"] ?? "Kingdee.BOS.WebApi.ServicesStub.AuthService.ValidateUser.common.kdsvc",
                Username = configuration["ERP:Username"] ?? "Administrator",
                Password = configuration["ERP:Password"] ?? "Ynk@407409",
                AcctID = configuration["ERP:AcctID"] ?? "674b4b07b00d63",
                Lcid = int.Parse(configuration["ERP:Lcid"] ?? "2052"),
                TimeoutSeconds = int.Parse(configuration["ERP:TimeoutSeconds"] ?? "30")
            };

            // 配置HttpClient
            _httpClient.Timeout = TimeSpan.FromSeconds(_config.TimeoutSeconds);
            _httpClient.DefaultRequestHeaders.Add("Accept", "*/*");
        }

        #region ERP收料通知单
        public async Task<ApiResponse> ProcessLotNoticeAsync(List<LotNoticeRequestYNK> requests)
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

                    if (request.ITEMID.StartsWith("Z") || request.ITEMID.StartsWith("H"))
                    {
                        _logger.LogWarning($"过滤Z和H开头的物料: {request.ITEMID}");
                        continue;
                    }
                    //判断重复
                    bool isExist = Db.Ado.GetInt($@"SELECT count(*) FROM INSPECT_IQC WHERE KEEID = '{request.KEEID}'") > 0;
                    if (isExist)
                    {
                        _logger.LogWarning($"收料通知单已存在: {request.KEEID}");
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

        public void SaveToDatabase(LotNoticeRequestYNK request, string inspectionId)
        {
            // 保存数据
            SaveMainInspection(request, inspectionId);
        }

        private void SaveMainInspection(LotNoticeRequestYNK request, string inspectionId)
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
            //string ItemID = Db.Ado.GetScalar($@"SELECT TOP 1 ITEMID FROM ITEM WHERE ITEMNAME = '{request.ITEMNAME}'")?.ToString().Trim();

            //if (string.IsNullOrEmpty(ItemID))
            //{
            //    ItemID = Db.Ado.GetScalar($@"select TOP 1 cast(cast(dbo.getNumericValue(ITEMID) AS DECIMAL)+1 as char) from ITEM order by ITEMID desc")?.ToString().Trim();
            //    if (string.IsNullOrEmpty(ItemID))
            //    {
            //        ItemID = "1001";
            //    }
            //    Db.Ado.ExecuteCommand($@"INSERT INTO ITEM (
            //                                TENID, ITEMID, ITEM0A17, ITEMCREATEUSER, ITEMCREATEDATE,
            //                                ITEMMODIFYDATE, ITEMMODIFYUSER, ITEMCODE, ITEMNAME)
            //                            VALUES (
            //                                '001', '{ItemID}', '001', 'system', '{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}',
            //                                '{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")}', 'system', '{ItemID}', '{request.ITEMNAME}')");
            //}

            Db.Ado.ExecuteCommand($@"
                MERGE INTO ITEM AS target
                USING (VALUES ('001', @ITEMID, '001', 'system', @CreateDate, 
                                @ModifyDate, 'system', @ITEMID, @ITEMNAME)) 
                        AS source (TENID, ITEMID, ITEM0A17, ITEMCREATEUSER, ITEMCREATEDATE,
                                    ITEMMODIFYDATE, ITEMMODIFYUSER, ITEMCODE, ITEMNAME)
                ON target.ITEMID = source.ITEMID AND target.TENID = source.TENID
                WHEN MATCHED THEN
                    UPDATE SET 
                        ITEM0A17 = source.ITEM0A17,
                        ITEMMODIFYDATE = source.ITEMMODIFYDATE,
                        ITEMMODIFYUSER = source.ITEMMODIFYUSER,
                        ITEMCODE = source.ITEMCODE,
                        ITEMNAME = source.ITEMNAME
                WHEN NOT MATCHED THEN
                    INSERT (TENID, ITEMID, ITEM0A17, ITEMCREATEUSER, ITEMCREATEDATE,
                            ITEMMODIFYDATE, ITEMMODIFYUSER, ITEMCODE, ITEMNAME)
                    VALUES (source.TENID, source.ITEMID, source.ITEM0A17, source.ITEMCREATEUSER, 
                            source.ITEMCREATEDATE, source.ITEMMODIFYDATE, source.ITEMMODIFYUSER, 
                            source.ITEMCODE, source.ITEMNAME);",
            new
            {
                ITEMID = request.ITEMID,
                ITEMNAME = request.ITEMNAME,
                CreateDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"),
                ModifyDate = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")
            });

            //更新INSPECT_IQC
            string sql = @"
                INSERT INTO INSPECT_IQC (
                    TENID, INSPECT_IQCID, INSPECT_IQCCREATEUSER, 
                    INSPECT_IQCCREATEDATE, ITEMNAME, ERP_ARRIVEDID, 
                    LOT_QTY, INSPECT_IQCCODE, ITEMID, LOTNO, 
                    APPLY_DATE, ITEM_SPECIFICATION, UNIT, 
                    SUPPID, SUPPNAME, SUPPLOTNO, KEEID, FID
                ) VALUES (
                    @TenId, @InspectIqcId, @InspectIqcCreateUser, 
                    getdate(), @ItemName, @ErpArrivedId,
                    @LotQty, @InspectIqcCode, @ItemId, @LotNo, 
                    @ApplyDate, @ItemSpecification, @Unit, 
                    @SuppID, @SuppName, @SuppLotNo, @KeeId, @FId
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
                new SugarParameter("@ItemId", request.ITEMID),
                new SugarParameter("@LotNo", (request.LOTNO==null?"":request.LOTNO.ToString())),
                new SugarParameter("@ApplyDate", request.APPLY_DATE),
                new SugarParameter("@ItemSpecification", request.MODEL_SPEC),
                new SugarParameter("@Unit", request.UNIT),
                new SugarParameter("@SuppID", SuppID),
                new SugarParameter("@SuppName", request.SUPPNAME),
                new SugarParameter("@SuppLotNo", request.SUPPLOTNO),
                new SugarParameter("@KeeId", request.KEEID),
                new SugarParameter("@FId", request.FID)
            };

            // 执行 SQL 命令
            Db.Ado.ExecuteCommand(sql, parameters);
        }
        #endregion

        #region 金蝶云登录接口,获取KDSVCSessionId 的值
        public async Task<ERPLoginResponseYNK> LoginAsync()
        {
            return await LoginAsync(_config.Username, _config.Password, _config.AcctID, _config.Lcid);
        }

        public async Task<ERPLoginResponseYNK> LoginAsync(string username, string password, string acctID, int lcid = 2052)
        {
            var response = new ERPLoginResponseYNK();

            try
            {
                // 创建请求参数
                var requestData = new ERPLoginRequestYNK
                {
                    Username = username,
                    Password = password,
                    AcctID = acctID,
                    Lcid = lcid
                };

                // 序列化为JSON
                var jsonContent = JsonConvert.SerializeObject(requestData);
                var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                // 构建完整的URL
                var url = $"{_config.BaseUrl}{_config.LoginUrl}";

                // 发送POST请求
                var httpResponse = await _httpClient.PostAsync(url, content);
                response.StatusCode = (int)httpResponse.StatusCode;

                // 检查响应状态
                if (httpResponse.IsSuccessStatusCode)
                {
                    var responseContent = await httpResponse.Content.ReadAsStringAsync();

                    // 从响应头中获取KDSVCSessionId
                    if (httpResponse.Headers.TryGetValues("Set-Cookie", out var cookies))
                    {
                        foreach (var cookie in cookies)
                        {
                            if (cookie.Contains("KDSVCSessionId"))
                            {
                                response.KDSVCSessionId = ExtractSessionIdFromCookie(cookie);
                                response.IsSuccess = true;
                                break;
                            }
                        }
                    }

                    // 如果没有在cookie中找到，尝试从响应体中解析
                    if (string.IsNullOrEmpty(response.KDSVCSessionId))
                    {
                        response.KDSVCSessionId = ParseSessionIdFromResponse(responseContent);
                        response.IsSuccess = !string.IsNullOrEmpty(response.KDSVCSessionId);
                    }

                    if (!response.IsSuccess)
                    {
                        response.ErrorMessage = "登录成功但未找到KDSVCSessionId";
                    }
                }
                else
                {
                    response.ErrorMessage = $"HTTP错误: {httpResponse.StatusCode}";
                    var errorContent = await httpResponse.Content.ReadAsStringAsync();
                    if (!string.IsNullOrEmpty(errorContent))
                    {
                        response.ErrorMessage += $", 错误信息: {errorContent}";
                    }
                }
            }
            catch (Exception ex)
            {
                response.ErrorMessage = $"登录失败: {ex.Message}";
                response.StatusCode = 500;
            }

            return response;
        }

        private string ExtractSessionIdFromCookie(string cookieHeader)
        {
            // 从Set-Cookie头中提取KDSVCSessionId
            var cookies = cookieHeader.Split(';');
            foreach (var cookie in cookies)
            {
                var trimmedCookie = cookie.Trim();
                if (trimmedCookie.StartsWith("KDSVCSessionId="))
                {
                    return trimmedCookie.Substring("KDSVCSessionId=".Length);
                }
            }
            return null;
        }

        private string ParseSessionIdFromResponse(string responseContent)
        {
            // 根据金蝶云API的实际响应格式解析SessionId
            try
            {
                // 假设响应格式为: {"LoginResultType":1,"Message":"","KDSVCSessionId":"session-id-here"}
                dynamic responseObj = JsonConvert.DeserializeObject(responseContent);
                return responseObj?.KDSVCSessionId?.ToString();
            }
            catch
            {
                return null;
            }
        }
        #endregion

        #region 回写ERP MES方法
        public List<LotNoticeResultRequestYNK> GetQmsLotNoticeResultRequest()
        {
            var sql = @"SELECT 
                            ERP_ARRIVEDID AS ERP_ARRIVEDID, 
                            INSPECT_IQCCODE AS INSPECT_IQCCODE,
                            ITEMID AS ITEMID,
                            ITEMNAME AS ITEMNAME,
                            LOTNO AS LOTNO,
                            FID AS FID,
                            UNIT,SUPPNAME,
                            KEEID AS FEntryID,
                            CASE 
                                WHEN COALESCE(SQM_STATE, OQC_STATE) IN ('OQC_STATE_005', 'OQC_STATE_006', 'OQC_STATE_008', 'OQC_STATE_011') THEN '合格'
                                WHEN COALESCE(SQM_STATE, OQC_STATE) = 'OQC_STATE_007' THEN '不合格'
                                WHEN COALESCE(SQM_STATE, OQC_STATE) = 'OQC_STATE_010' THEN '免检'
                                ELSE '不合格'
                            END AS OQC_STATE,
                            ISNULL(TRY_CAST(FQC_CNT AS DECIMAL), 0) AS FReceiveQty,       -- 不可转换时返回0
                            ISNULL(TRY_CAST(FQC_NOT_CNT AS DECIMAL), 0) AS FRefuseQty     -- 不可转换时返回0
                        FROM INSPECT_IQC
                        WHERE (ISSY <> '1' OR ISSY IS NULL) 
                            AND COALESCE(SQM_STATE, OQC_STATE) IN ('OQC_STATE_005', 'OQC_STATE_006', 'OQC_STATE_007', 'OQC_STATE_008', 'OQC_STATE_010', 'OQC_STATE_011')
                        ORDER BY INSPECT_IQCCREATEDATE DESC;";

            var list = Db.Ado.SqlQuery<LotNoticeResultRequestYNK>(sql);
            foreach (var item in list)
            {
                item.FCheckQty = Db.Ado.GetDecimal(@$"SELECT COALESCE((
    SELECT TOP 1 
        MAX(CAST(INSPECT_CNT AS INT)) OVER (PARTITION BY INSPECT_TYPE) AS INSPECT_CNT
    FROM INSPECT_PROGRESS 
    WHERE DOC_CODE = '{item.INSPECT_IQCCODE}'
        AND INSPECT_TYPE IN ('INSPECT_TYPE_002', 'INSPECT_TYPE_001', 'INSPECT_TYPE_003', 'INSPECT_TYPE_999', 'INSPECT_TYPE_000')
    ORDER BY 
        CASE INSPECT_TYPE 
            WHEN 'INSPECT_TYPE_002' THEN 1
            WHEN 'INSPECT_TYPE_001' THEN 2
            WHEN 'INSPECT_TYPE_003' THEN 3
            WHEN 'INSPECT_TYPE_999' THEN 4
            WHEN 'INSPECT_TYPE_000' THEN 5
        END
), 0) AS INSPECT_CNT");
            }
            return list;
        }

        public void CallBackQmsLotNoticeResult(LotNoticeResultRequestYNK request)
        {
            var sql = string.Format(@"update INSPECT_IQC set ISSY='1' where KEEID='{0}' ", request.FEntryID);
            Db.Ado.ExecuteCommand(sql);
        }

        public List<AttachmentResultRequestYNK> GetAttachmentResultRequest()
        {
            var sql = @"SELECT TOP 10 SCANDOC.SCANDOCID,
                            ERP_ARRIVEDID AS ERP_ARRIVEDID,
                            INSPECT_IQCCODE AS INSPECT_IQCCODE,
                            FID AS FID,
                            KEEID AS FEntryID
                        FROM SCANDOC 
                            LEFT JOIN INSPECT_IQC IQC ON INSPECT_IQCID=SCANDOC.PEOPLEID
                            WHERE (SCANDOC.ISSY <> '1' OR SCANDOC.ISSY IS NULL) 
                            AND COALESCE(SQM_STATE, OQC_STATE) IN ('OQC_STATE_005', 'OQC_STATE_006', 'OQC_STATE_007', 'OQC_STATE_008', 'OQC_STATE_010', 'OQC_STATE_011')
                        ORDER BY INSPECT_IQCCREATEDATE DESC;";

            var list = Db.Ado.SqlQuery<AttachmentResultRequestYNK>(sql);
            foreach (var item in list)
            {
                string filePath = Path.Combine(AppSettings.Configuration["AppSettings:FileServerPath"], item.FilePath.TrimStart('\\'));
                if (File.Exists(filePath))
                {
                    item.SendBytes = File.ReadAllBytes(filePath);
                }
                else
                {
                    throw new Exception("找不到文件:" + filePath);
                }
            }
            return list;
        }

        public void CallBackAttachmentResult(AttachmentResultRequestYNK request)
        {
            var sql = string.Format(@"update SCANDOC set ISSY='1' where SCANDOC='{0}' ", request.SCANDOCID);
            Db.Ado.ExecuteCommand(sql);
        }

        #endregion

        #region 工具API
        public async Task<ApiResponse> GetAOIInspectInfoByDocCodeAsync(INSPECT_REQCODE input)
        {
            try
            {
                var parameters = new SugarParameter[] {
                  new SugarParameter("@DOC_CODE", input.DOC_CODE) };

                var data = await Db.Ado.SqlQueryAsync<INSPECT_INFO_BYCODE>(
                    "select TOP 1 ITEMID,ITEMNAME,LOTNO,LOT_QTY from INSPECT_VIEW where INSPECT_CODE = @DOC_CODE",
                    parameters
                );

                return new ApiResponse
                {
                    Success = true,
                    Message = "数据获取成功",
                    Data = JsonConvert.SerializeObject(data.FirstOrDefault())
                };
            }
            catch (Exception ex)
            {
                return new ApiResponse
                {
                    Success = false,
                    Message = $"数据获取失败：{ex.Message}"
                };
            }
        }

        public async Task<ApiResponse> GetAOIProgressDataByDocCodeAsync(INSPECT_REQCODE input)
        {
            try
            {
                var parameters = new SugarParameter[] {
                  new SugarParameter("@DOC_CODE", input.DOC_CODE) };

                List<INSPECT_PROGRESS_BYCODE> data = await Db.Ado.SqlQueryAsync<INSPECT_PROGRESS_BYCODE>(
                    "select INSPECT_PROGRESSNAME, INSPECT_PROGRESSID from INSPECT_PROGRESS where DOC_CODE = @DOC_CODE and INSPECT_DEV='INSPECT_DEV_010'",
                    parameters
                );

                return new ApiResponse
                {
                    Success = true,
                    Message = "数据获取成功",
                    Data = JsonConvert.SerializeObject(data)
                };
            }
            catch (Exception ex)
            {
                return new ApiResponse
                {
                    Success = false,
                    Message = $"数据获取失败：{ex.Message}"
                };
            }
        }

        public async Task<ApiResponse> ProcessUploadAOIDataAsync(List<InspectAoi> input)
        {
            try
            {
                foreach (var item in input)
                {
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
                    if (!exists)
                    {
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

                return new ApiResponse
                {
                    Success = true,
                    Message = "Aoi数据保存成功"
                };
            }
            catch (Exception ex)
            {
                return new ApiResponse
                {
                    Success = false,
                    Message = $"Aoi数据保存失败：{ex.Message}"
                };
            }
        }

        public async Task<ApiResponse> ProcessUploadAOIImageDataAsync(List<InspectImageAoi> input)
        {
            try
            {
                if (input.Count == 0)
                {
                    return new ApiResponse
                    {
                        Success = false,
                        Message = $"传入数据为空!"
                    };
                }
                string baseDirPath = Path.Combine(AppSettings.Configuration["AppSettings:FileServerPath"], @"AOI");
                foreach (var item in input)
                {
                    // 1. 验证必要数据
                    if (string.IsNullOrEmpty(item.DOC_CODE))
                    {
                        throw new ArgumentException("DOC_CODE不能为空，无法存储图片");
                    }
                    if (string.IsNullOrEmpty(item.ImageName))
                    {
                        throw new ArgumentException("ImageName不能为空，无法确定文件名");
                    }
                    if (string.IsNullOrEmpty(item.ImageData))
                    {
                        throw new ArgumentException($"DOC_CODE: {item.DOC_CODE} 的图片数据为空，无法存储");
                    }

                    // 2. 创建DOC_CODE对应的文件夹
                    string docCodeDirPath = Path.Combine(baseDirPath, item.DOC_CODE);
                    if (!Directory.Exists(docCodeDirPath))
                    {
                        Directory.CreateDirectory(docCodeDirPath);
                    }

                    // 3. 处理文件名（存在时添加_1、_2等后缀）
                    string fileExtension = Path.GetExtension(item.ImageName);
                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(item.ImageName);
                    string targetFilePath = Path.Combine(docCodeDirPath, item.ImageName);

                    // 检查文件是否存在，存在则生成新文件名
                    if (File.Exists(targetFilePath))
                    {
                        int counter = 1;
                        do
                        {
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

                return new ApiResponse
                {
                    Success = true,
                    Message = "AOI图片保存成功",
                    Data = JsonConvert.SerializeObject(input)
                };
            }
            catch (Exception ex)
            {
                return new ApiResponse
                {
                    Success = false,
                    Message = $"AOI图片数据保存失败：{ex.Message}"
                };
            }
        }

        public async Task<ApiResponse> ProcessYNKInpectProcessDataAsync(List<INSPECT_PROGRESSDto> input)
        {
            try
            {
                if (input.Count == 0)
                {
                    return new ApiResponse
                    {
                        Success = false,
                        Message = $"传入数据为空!"
                    };
                }
                var firstEntity = input[0];
                var parameters = new SugarParameter[] {
                  new SugarParameter("@DOC_CODE", firstEntity.DOC_CODE) };

                // 1. 查询历史数据（包含版本、检验项目和顺序号）
                var dtOldData = Db.Ado.GetDataTable(
                    @"select VER,INSPECT_PROGRESSNAME,OID,INSPECT_CNT,INSPECT_PLANID
      from INSPECT_PROGRESS 
      where DOC_CODE = @DOC_CODE and COC_ATTR='COC_ATTR_001'",
                    parameters
                );

                var dtEnum = dtOldData.AsEnumerable();
                string newVer = "01"; // 新版本号
                int lastMaxVer = 0;   // 上一版本号（数字形式）

                // 2. 处理版本号逻辑
                if (dtOldData.Rows.Count > 0)
                {
                    lastMaxVer = dtEnum
                        .Select(row => {
                            int.TryParse(row["VER"].ToString().TrimStart('0'), out int v);
                            return v;
                        })
                        .Max();
                    newVer = (lastMaxVer + 1).ToString("00");
                }
                if (newVer == "01")
                {
                    int oIdIndex = 1;
                    int INSPECT_CNT = 0;
                    var entityType = firstEntity.GetType();
                    // 遍历A1到A64的所有属性
                    for (int i = 1; i <= 64; i++)
                    {
                        // 构造属性名（A1, A2, ..., A64）
                        string propertyName = $"A{i}";
                        // 获取属性信息
                        var property = entityType.GetProperty(propertyName);
                        if (property != null)
                        {
                            var value = property.GetValue(firstEntity);
                            if (value != null)
                            {
                                if (value is string strValue)
                                {
                                    if (!string.IsNullOrEmpty(strValue.Trim()))
                                    {
                                        INSPECT_CNT++;
                                    }
                                }
                            }
                        }
                    }
                    // 定义查询参数（避免 SQL 注入）
                    var planParameters = new SugarParameter[] {
                        new SugarParameter("@SPOT_CNT", INSPECT_CNT)};
                    // 查询 INSPECT_PLAN 表，获取 SPOT_CNT 等于样本数量的 INSPECT_PLANID
                    var planId = Db.Ado.GetString(
                        @"select INSPECT_PLANID from INSPECT_PLAN where SPOT_CNT = @SPOT_CNT", // 条件：样本数量匹配
                        planParameters
                    );
                    foreach (var item in input)
                    {
                        item.VER = newVer; // 设置新版本号
                        item.OID = (oIdIndex++).ToString("00");
                        item.INSPECT_CNT = INSPECT_CNT.ToString();
                        item.INSPECT_PLANID = planId;
                    }
                }
                else
                {//第二次上传
                 // 非首次上传：处理OID逻辑
                 // 2.1 提取历史数据中每个检验项目最近出现的OID（按版本倒序取最近）
                    var lastestOidMap = dtEnum
                        .GroupBy(row => row["INSPECT_PROGRESSNAME"].ToString(), StringComparer.OrdinalIgnoreCase)
                        .ToDictionary(
                            group => group.Key,
                            group => {
                                // 按版本号降序排序，取第一个（最近版本）的OID
                                var latestRow = group
                                    .OrderByDescending(row => {
                                        int.TryParse(row["VER"].ToString().TrimStart('0'), out int v);
                                        return v;
                                    })
                                    .FirstOrDefault();

                                // 转换OID为整数（默认0）
                                if (latestRow != null && int.TryParse(latestRow["OID"].ToString(), out int oid))
                                {
                                    return oid;
                                }
                                return 0;
                            }
                        );

                    // 2.2 获取历史数据中最大的OID（用于新增项目累加）
                    int maxHistoryOid = dtEnum
                        .Select(row => {
                            int.TryParse(row["OID"].ToString(), out int oid);
                            return oid;
                        })
                        .DefaultIfEmpty(0)
                        .Max();
                    var firstVerRow = dtEnum
                        .FirstOrDefault();  // 取第一条符合条件的记录

                    // 2.3 遍历输入项分配OID
                    int currentMaxOid = maxHistoryOid; // 当前最大OID（用于累加）
                    foreach (var item in input)
                    {
                        item.VER = newVer;
                        item.INSPECT_CNT = firstVerRow["INSPECT_CNT"].ToString();
                        item.INSPECT_PLANID = firstVerRow["INSPECT_PLANID"].ToString();
                        // 检查当前检验项目是否在历史记录中存在
                        if (lastestOidMap.TryGetValue(item.INSPECT_PROGRESSNAME, out int existOid) && existOid > 0)
                        {
                            // 规则2：存在则使用最近版本的OID
                            item.OID = existOid.ToString("00");
                        }
                        else
                        {
                            // 规则1：不存在则从最大OID累加
                            currentMaxOid++;
                            item.OID = currentMaxOid.ToString("00");
                        }
                    }
                }

                #region 保存数据
                await SaveInspectProgressList(input);
                #endregion

                return new ApiResponse
                {
                    Success = true,
                    Message = "数据保存成功",
                    Data = JsonConvert.SerializeObject(input)
                };
            }
            catch (Exception ex)
            {
                return new ApiResponse
                {
                    Success = false,
                    Message = $"二次元数据保存失败：{ex.Message}"
                };
            }
        }

        /// <summary>
        /// 批量保存检验进度数据
        /// </summary>
        /// <param name="input">检验进度实体数组</param>
        /// <returns>是否保存成功</returns>
        private async Task SaveInspectProgressList(List<INSPECT_PROGRESSDto> input)
        {
            if (input == null || input.Count == 0)
                return;

            // 每个实体需要的基础参数数量：18个基础字段 + 64个A字段 = 82个
            // 但通过复用相同值的参数，实际数量会减少
            int parametersPerItem = 82;
            int maxBatchSize = 2100 / parametersPerItem; // 仍保持分批处理基础逻辑

            for (int i = 0; i < input.Count; i += maxBatchSize)
            {
                var batchItems = input.Skip(i).Take(maxBatchSize).ToList();
                if (batchItems.Count == 0)
                    continue;

                var (sql, parameters) = BuildBatchSqlWithReusedParameters(batchItems);
                await Db.Ado.ExecuteCommandAsync(sql, parameters.ToArray());
            }
        }

        private (string Sql, List<SugarParameter> Parameters) BuildBatchSqlWithReusedParameters(List<INSPECT_PROGRESSDto> batchItems)
        {
            var sqlBuilder = new StringBuilder();
            sqlBuilder.Append("INSERT INTO INSPECT_PROGRESS (");
            sqlBuilder.Append("INSPECT_PROGRESSID, DOC_CODE, ITEMID,INSPECT02CODE, VER, OID, COC_ATTR, ");
            sqlBuilder.Append("INSPECT_PROGRESSNAME, INSPECT_DEV, COUNTTYPE, INSPECT_PLANID, ");
            sqlBuilder.Append("INSPECT_CNT, STD_VALUE, MAX_VALUE, MIN_VALUE, UP_VALUE, DOWN_VALUE, ");

            // 拼接A1-A64样本字段
            for (int i = 1; i <= 64; i++)
            {
                sqlBuilder.Append($"A{i}, ");
            }

            sqlBuilder.Append("INSPECT_PROGRESSCREATEUSER, INSPECT_PROGRESSCREATEDATE, TENID");
            sqlBuilder.Append(") VALUES ");

            var parameters = new List<SugarParameter>();
            var parameterCache = new Dictionary<string, string>(); // 缓存值与参数名的映射
            int paramIndex = 0;

            foreach (var item in batchItems)
            {
                sqlBuilder.Append("(");

                // 处理主键ID（通常唯一，难以复用）
                paramIndex++;
                var progressIdParamName = $"@INSPECT_PROGRESSID_{paramIndex}";
                sqlBuilder.Append($"{progressIdParamName}, ");
                parameters.Add(new SugarParameter(progressIdParamName,
                    string.IsNullOrEmpty(item.INSPECT_PROGRESSID) ? Guid.NewGuid().ToString() : item.INSPECT_PROGRESSID));

                // 处理可复用的字段 - 使用值作为键缓存参数名
                sqlBuilder.Append(AddReusableParameter(
                    "DOC_CODE", item.DOC_CODE, ref paramIndex, parameters, parameterCache) + ", ");

                sqlBuilder.Append(AddReusableParameter(
                    "ITEMID", item.ITEMID, ref paramIndex, parameters, parameterCache) + ", ");

                sqlBuilder.Append(AddReusableParameter(
                   "INSPECT02CODE", item.INSPECT02CODE, ref paramIndex, parameters, parameterCache) + ", ");

                sqlBuilder.Append(AddReusableParameter(
                    "VER", item.VER, ref paramIndex, parameters, parameterCache) + ", ");

                sqlBuilder.Append(AddReusableParameter(
                    "OID", item.OID, ref paramIndex, parameters, parameterCache) + ", ");

                sqlBuilder.Append(AddReusableParameter(
                    "COC_ATTR", item.COC_ATTR, ref paramIndex, parameters, parameterCache) + ", ");

                sqlBuilder.Append(AddReusableParameter(
                    "INSPECT_PROGRESSNAME", item.INSPECT_PROGRESSNAME, ref paramIndex, parameters, parameterCache) + ", ");

                sqlBuilder.Append(AddReusableParameter(
                    "INSPECT_DEV", item.INSPECT_DEV, ref paramIndex, parameters, parameterCache) + ", ");

                sqlBuilder.Append(AddReusableParameter(
                    "COUNTTYPE", item.COUNTTYPE, ref paramIndex, parameters, parameterCache) + ", ");

                sqlBuilder.Append(AddReusableParameter(
                    "INSPECT_PLANID", item.INSPECT_PLANID, ref paramIndex, parameters, parameterCache) + ", ");

                sqlBuilder.Append(AddReusableParameter(
                    "INSPECT_CNT", item.INSPECT_CNT, ref paramIndex, parameters, parameterCache) + ", ");

                sqlBuilder.Append(AddReusableParameter(
                    "STD_VALUE", item.STD_VALUE, ref paramIndex, parameters, parameterCache) + ", ");

                sqlBuilder.Append(AddReusableParameter(
                    "MAX_VALUE", item.MAX_VALUE, ref paramIndex, parameters, parameterCache) + ", ");

                sqlBuilder.Append(AddReusableParameter(
                    "MIN_VALUE", item.MIN_VALUE, ref paramIndex, parameters, parameterCache) + ", ");

                sqlBuilder.Append(AddReusableParameter(
                    "UP_VALUE", item.UP_VALUE, ref paramIndex, parameters, parameterCache) + ", ");

                sqlBuilder.Append(AddReusableParameter(
                    "DOWN_VALUE", item.DOWN_VALUE, ref paramIndex, parameters, parameterCache) + ", ");

                // 处理A1-A64样本字段
                for (int j = 1; j <= 64; j++)
                {
                    var propName = $"A{j}";
                    var propValue = item.GetType().GetProperty(propName)?.GetValue(item);
                    var paramKey = $"{propName}_{propValue}";

                    sqlBuilder.Append(AddReusableParameter(
                        propName, propValue, ref paramIndex, parameters, parameterCache) + (j < 64 ? ", " : ""));
                }

                // 处理创建用户和日期
                sqlBuilder.Append(", " + AddReusableParameter(
                    "INSPECT_PROGRESSCREATEUSER", item.INSPECT_PROGRESSCREATEUSER, ref paramIndex, parameters, parameterCache));

                sqlBuilder.Append(", " + AddReusableParameter(
                    "INSPECT_PROGRESSCREATEDATE", item.INSPECT_PROGRESSDATE, ref paramIndex, parameters, parameterCache));

                sqlBuilder.Append(", " + AddReusableParameter(
                    "TENID", item.TENID, ref paramIndex, parameters, parameterCache));

                sqlBuilder.Append("),");
            }

            // 移除最后一个逗号
            if (sqlBuilder.Length > 0 && sqlBuilder[sqlBuilder.Length - 1] == ',')
            {
                sqlBuilder.Length--;
            }

            return (sqlBuilder.ToString(), parameters);
        }

        // 复用参数的核心方法：相同值使用同一个参数
        private string AddReusableParameter(string fieldName, object value, ref int paramIndex,
            List<SugarParameter> parameters, Dictionary<string, string> parameterCache)
        {
            // 创建唯一键：字段名+值（处理null情况）
            var cacheKey = $"{fieldName}_{(value ?? "NULL").ToString()}";

            // 如果已有相同值的参数，直接返回已存在的参数名
            if (parameterCache.TryGetValue(cacheKey, out var existingParamName))
            {
                return existingParamName;
            }

            // 否则创建新参数
            paramIndex++;
            var newParamName = $"@{fieldName}_{paramIndex}";
            parameters.Add(new SugarParameter(newParamName, value ?? DBNull.Value));
            parameterCache[cacheKey] = newParamName;

            return newParamName;
        }
        #endregion

        #region 报表相关
        public async Task<ApiResponse> GetInspectionRecordReportDataAsync(INSPECT_REQCODE input)
        {
            try
            {
                var parameters = new SugarParameter[] {
                  new SugarParameter("@DOC_CODE", input.DOC_CODE) };

                string docNo = string.Empty;
                string docCreateDate = string.Empty;
                string orgMoa02 = string.Empty;
                var data1 = await Db.Ado.GetDataTableAsync("select TOP 1 DOC_NO,DOC_CREATEDATE from INSPECT_DOC_INFO where IS_NEW='1' and DOC_TYPE='图纸尺寸文件'");
                if (data1 != null && data1.Rows.Count > 0)
                {
                    docNo = data1.Rows[0]["DOC_NO"].ToString();
                    docCreateDate = data1.Rows[0]["DOC_CREATEDATE"].ToString();
                }
                var data2 = await Db.Ado.GetDataTableAsync("select TOP 1 ORGM0A02 from ORGM001 where ORGM0A01='001'");
                if (data2 != null && data2.Rows.Count > 0)
                {
                    orgMoa02 = data2.Rows[0]["ORGM0A02"].ToString();
                }
                var dataMain = await Db.Ado.GetDataTableAsync(@"SELECT TOP 1 ITEMID,ITEMNAME,LOTNO,LOT_QTY,INSPECT_IQCNAME,
        CASE 
            WHEN COALESCE(SQM_STATE, OQC_STATE) IN ('OQC_STATE_005', 'OQC_STATE_006', 'OQC_STATE_008', 'OQC_STATE_011') THEN '合格'
            WHEN COALESCE(SQM_STATE, OQC_STATE) = 'OQC_STATE_007' THEN '不合格'
            WHEN COALESCE(SQM_STATE, OQC_STATE) = 'OQC_STATE_010' THEN '免检'
            ELSE OQC_STATE
        END AS OQC_STATE
    FROM INSPECT_IQC
    WHERE INSPECT_IQCCODE=@DOC_CODE", parameters);
                string ITEMID = string.Empty;
                string ITEMNAME = string.Empty;
                string LOTNO = string.Empty;
                string LOT_QTY = string.Empty;
                string OQC_STATE = string.Empty;
                string Inspector = string.Empty;
                string InspectorDate = string.Empty;
                if (dataMain != null && dataMain.Rows.Count > 0)
                {
                    ITEMID = dataMain.Rows[0]["ITEMID"].ToString();
                    ITEMNAME = dataMain.Rows[0]["ITEMNAME"].ToString();
                    LOTNO = dataMain.Rows[0]["LOTNO"].ToString();
                    LOT_QTY = dataMain.Rows[0]["LOT_QTY"].ToString();
                    OQC_STATE = dataMain.Rows[0]["OQC_STATE"].ToString();
                    InspectorDate = dataMain.Rows[0]["INSPECT_IQCNAME"].ToString();
                }
                var originalData = await Db.Ado.GetDataTableAsync(@"SELECT CASE 
        WHEN STD_VALUE IS NULL OR STD_VALUE='' THEN 'N/A'
        ELSE STD_VALUE + '+' + MAX_VALUE + '/-' + replace(MIN_VALUE,'-','')
    END AS  PROGRESSNAME1,
    CASE 
        WHEN INSPECT_TYPE='INSPECT_TYPE_002' THEN '1'
        ELSE '2'
    END AS  INSPECT_TYPE1,
    CASE 
        WHEN INSPECT02CODE IS NULL OR LTRIM(RTRIM(INSPECT02CODE)) = '' 
        THEN SYSM002.SYSM002NAME
        ELSE SYSM002.SYSM002NAME + '<br>' + INSPECT02CODE 
    END AS INSPECT02CODE, NGS, COUNTTYPE,USER_.USER_NAME AS INSPECTOR,
            (case when INSPECT_RESULT='INSPECT_RESULT_001' then 'OK' when INSPECT_RESULT='INSPECT_RESULT_002' then 'NG' else '' end) INSPECT_RESULT, 
            A1, A2, A3, A4, A5, A6, A7, A8, A9, A10, A11, A12, A13, A14, A15, A16, 
            A17, A18, A19, A20, A21, A22, A23, A24, A25, A26, A27, A28, A29, A30, A31, A32, 
            A33, A34, A35, A36, A37, A38, A39, A40, A41, A42, A43, A44, A45, A46, A47, A48, 
            A49, A50, A51, A52, A53, A54, A55, A56, A57, A58, A59, A60, A61, A62, A63, A64
            FROM INSPECT_PROGRESS
			LEFT JOIN SYSM002 on SYSM002.SYSM002ID=INSPECT_DEV
            LEFT JOIN USER_ on USER_.User_Account=INSPECT_PROGRESS.INSPECT_PROGRESSMODIFYUSER
            WHERE DOC_CODE = @DOC_CODE and INSPECT_TYPE in ('INSPECT_TYPE_002','INSPECT_TYPE_000','INSPECT_TYPE_003','INSPECT_TYPE_999') 
            ORDER BY INSPECT_TYPE1,INSPECT_PROGRESSNAME"
                , parameters);//and COC_ATTR='COC_ATTR_001'

                // 将原始数据转换为对象列表
                var dataList = new List<InspectData>();

                foreach (DataRow row in originalData.Rows)
                {
                    Inspector = row["INSPECTOR"].ToString();
                    var data = new InspectData
                    {
                        ProgressName = row["PROGRESSNAME1"]?.ToString(),
                        InspectCode = row["INSPECT02CODE"]?.ToString(),
                        InspectrResult = row["INSPECT_RESULT"]?.ToString(),
                        InspectrType = row["COUNTTYPE"]?.ToString(),
                        NGS = row["NGS"]?.ToString(),
                        Values = new List<object>()
                    };

                    // 收集A1-A64的值（跳过空值）
                    for (int i = 1; i <= 64; i++)
                    {
                        string fieldName = $"A{i}";
                        object value = row[fieldName];
                        if (value != null && !string.IsNullOrEmpty(value.ToString()))
                        {
                            data.Values.Add(value);
                        }
                    }

                    dataList.Add(data);
                }
                int pageRowCount = 13;
                int pageColCount = 13;
                // 分组逻辑
                var groupedData = GroupData(dataList, pageColCount, pageRowCount);

                // 构建最终的JSON结构
                var result = new
                {
                    PageRowCount = pageRowCount,
                    FormCode = docNo,
                    RecordNo = input.DOC_CODE,
                    MaterialCode = ITEMID,
                    MaterialName = ITEMNAME,
                    BatchNo = LOTNO,
                    Qty = LOT_QTY,
                    EffectiveDate = docCreateDate,
                    InspectionResult = OQC_STATE,
                    Remark = "",
                    Inspector = Inspector,
                    InspectorDate = InspectorDate,
                    Reviewer = "",
                    ReviewerDate = "",
                    DataValues = groupedData
                };
                string jsonData = JsonConvert.SerializeObject(result);

                return new ApiResponse
                {
                    Success = true,
                    Message = "数据获取成功",
                    Data = jsonData
                };
            }
            catch (Exception ex)
            {
                return new ApiResponse
                {
                    Success = false,
                    Message = $"数据获取失败：{ex.Message}"
                };
            }
        }

        // 数据分组方法
        private List<Dictionary<string, List<object>>> GroupData(List<InspectData> dataList, int pageColCount, int pageRowCount)
        {
            var result = new List<Dictionary<string, List<object>>>();

            if (dataList == null || dataList.Count == 0)
                return result;

            int maxLength = dataList.Select(d => d.Values?.Count ?? 0).Max();
            if (maxLength < 3) maxLength = 3; // 确保至少有3个基础元素

            int columnIndex = 1;
            // 按页分组处理数据
            for (int pageIndex = 0; pageIndex < dataList.Count; pageIndex += pageColCount)
            {
                var group = new Dictionary<string, List<object>>();

                // 创建 Column0 - 序号列
                var column0 = new List<object>();
                for (int j = 1; j <= maxLength; j++)
                {
                    column0.Add(j);
                }
                group.Add("Column0", column0);

                // 处理当前页的每一列数据
                for (int i = pageIndex; i < pageIndex + pageColCount && i < dataList.Count; i++)
                {
                    var data = dataList[i];
                    var columnData = new List<object>();

                    // 添加基础信息：序号、检测名称、检测编码
                    columnData.Add(columnIndex); // 序号
                    columnData.Add(data.ProgressName ?? "");  // 检测名称
                    columnData.Add(data.InspectCode ?? "");   // 检测编码

                    // 处理检测值
                    if (data.InspectrType == "COUNTTYPE_001")
                    {
                        // 计数类型处理
                        for (int j = 1; j <= maxLength; j++)
                        {
                            string resultValue = "OK";
                            if (string.IsNullOrEmpty(data.NGS))
                            {
                                resultValue = data.InspectrResult ?? "OK";
                            }
                            else
                            {
                                if (data.NGS.Contains($"A{j};"))
                                {
                                    resultValue = "NG";
                                }
                            }
                            columnData.Add(resultValue);
                        }
                    }
                    else
                    {
                        // 其他类型直接添加值
                        if (data.Values != null)
                        {
                            columnData.AddRange(data.Values);
                        }
                        // 填充剩余位置
                        while (columnData.Count < 3 + maxLength)
                        {
                            columnData.Add(""); // 空值填充
                        }
                    }

                    group.Add($"Column{columnIndex - pageIndex}", columnData);
                    columnIndex++;
                }

                // 创建 ColumnA - 结果汇总列
                var columnA = new List<object>();
                for (int i = 0; i < maxLength; i++)
                {
                    columnA.Add("OK"); // 默认结果
                }

                // 填充空列以达到页面列数要求
                while ((columnIndex - 1) % pageColCount != 0)
                {
                    var emptyColumn = new List<object> { columnIndex };
                    // 填充基础结构
                    for (int i = 1; i < 3 + maxLength; i++)
                    {
                        emptyColumn.Add(""); // 空值
                    }
                    group.Add($"Column{columnIndex - pageIndex}", emptyColumn);
                    columnIndex++;
                }
                group.Add("ColumnA", columnA);

                // 按行数拆分大组
                int totalRows = group.Values.Max(list => list?.Count ?? 0);
                int splitCount = (int)Math.Ceiling((double)(totalRows - 3) / pageRowCount);

                for (int splitIndex = 0; splitIndex < splitCount; splitIndex++)
                {
                    var splitGroup = new Dictionary<string, List<object>>();

                    foreach (var kvp in group)
                    {
                        var originalList = kvp.Value ?? new List<object>();
                        var splitList = new List<object>();

                        if (kvp.Key == "Column0" || kvp.Key == "ColumnA")
                        {
                            // 序号列和结果列：只拆分数据部分
                            if (splitIndex == 0 && originalList.Count > 0)
                            {
                                splitList.AddRange(originalList);
                            }
                            else
                            {
                                // 后续分组中，Column0和ColumnA只包含当前分页的数据
                                int startIndex = splitIndex * pageRowCount;
                                int takeCount = Math.Min(pageRowCount, originalList.Count - startIndex);
                                if (takeCount > 0)
                                {
                                    splitList.AddRange(originalList.GetRange(startIndex, takeCount));
                                }
                            }
                        }
                        else
                        {
                            // 数据列：前3个固定元素 + 当前分页的数据
                            // 添加固定元素（序号、名称、编码）
                            if (originalList.Count >= 1) splitList.Add(originalList[0]);
                            if (originalList.Count >= 2) splitList.Add(originalList[1]);
                            if (originalList.Count >= 3) splitList.Add(originalList[2]);

                            // 添加当前分页的数据
                            int dataStartIndex = 3 + splitIndex * pageRowCount;
                            int dataTakeCount = Math.Min(pageRowCount, originalList.Count - dataStartIndex);
                            if (dataTakeCount > 0)
                            {
                                splitList.AddRange(originalList.GetRange(dataStartIndex, dataTakeCount));
                            }
                        }

                        splitGroup.Add(kvp.Key, splitList);
                    }

                    result.Add(splitGroup);
                }
            }

            return result;
        }

        // 数据模型类
        public class InspectData
        {
            public string ProgressName { get; set; }
            public string InspectCode { get; set; }
            public string InspectrResult { get; set; }
            public string InspectrType { get; set; }
            public string NGS { get; set; }
            public List<object> Values { get; set; }
        }
        #endregion

        #region 看板相关
        private async Task<DataTable> GetPersonnelData(INSPECT_PERSONNELDATA input) {
            if (input == null) return null;
            string sql = @"SELECT INSPECT_IQCCREATEDATE,COMP_DATE,LOT_QTY,ITEMKIND,PROGRESS.INSPECTOR
from INSPECT_IQC 
INNER JOIN (SELECT DOC_CODE,max(USER_.USER_NAME) AS INSPECTOR FROM INSPECT_PROGRESS
LEFT JOIN USER_ on USER_.User_Account=INSPECT_PROGRESS.INSPECT_PROGRESSMODIFYUSER
WHERE DOC_CODE  is not null
GROUP by DOC_CODE having (max(USER_.USER_NAME) is not null 
and max(USER_.USER_NAME)!= '管理员' and max(USER_.USER_NAME)!= '管理员1')) PROGRESS ON PROGRESS.DOC_CODE=INSPECT_IQC.INSPECT_IQCCODE
where COMP_DATE is not null and ITEMKIND in ('塑胶件','金属件','电子件','辅料','包材','成品外购标准件')";

            if (input.SumType.ToLower() == "year") {
                sql += $" and year(INSPECT_IQCCREATEDATE) = '{DateTime.Today.Year}'";
            }
            else if (input.SumType.ToLower() == "month") {
                sql += $" and INSPECT_IQCCREATEDATE >= '{DateTime.Today.Year}-{DateTime.Today.Month}-01' and INSPECT_IQCCREATEDATE < '{DateTime.Today.AddMonths(1).Year}-{DateTime.Today.AddMonths(1).Month}-01'";
            }
            else {
                sql += $" and INSPECT_IQCCREATEDATE >= '{input.StartDate}' and INSPECT_IQCCREATEDATE < '{DateTime.Parse(input.EndDate).AddDays(1).ToString("yyyy-MM-dd")}'";
            }
            if (!string.IsNullOrEmpty(input.MeterialNames)) {
                var materialItems = input.MeterialNames.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                           .Select(item => item.Trim())
                                           .Where(item => !string.IsNullOrWhiteSpace(item));
                if (materialItems.Any()){ 
                    var materialParams = string.Join(",", materialItems.Select(item => $"'{item}'"));
                    sql += $" and ITEMKIND in ({materialParams})";
                }
            }
            var originalData = await Db.Ado.GetDataTableAsync(sql);
            return originalData;
        }

        /// <summary>
        /// 人员检验批数
        /// </summary>
        public async Task<ApiResponse> GetPersonnelBatchesDataAsync(INSPECT_PERSONNELDATA input)
        {
            try
            {
                //INSPECTOR,INSPECT_IQCCREATEDATE,LOT_QTY
                DataTable dt = await GetPersonnelData(input);
                // 转换为可枚举数据并处理空值
                var dataList = dt.AsEnumerable()
                    .Select(row => new {
                        Inspector = row.IsNull("INSPECTOR") ? string.Empty : row.Field<string>("INSPECTOR"),
                        CreateDate = row.IsNull("INSPECT_IQCCREATEDATE") ? DateTime.MinValue : row.Field<DateTime>("INSPECT_IQCCREATEDATE"),
                        LotQty = row.IsNull("LOT_QTY") ? 0 : Convert.ToDecimal(row.Field<object>("LOT_QTY"))
                    })
                    .Where(x => !string.IsNullOrEmpty(x.Inspector) && x.CreateDate != DateTime.MinValue)
                    .ToList();

                object result = null;

                if (input.SumType.ToLower() == "month") {
                    var monthStats = dataList
                        .GroupBy(x => x.Inspector)
                        .Select(g => new {
                            name = g.Key,
                            data = new[] { g.Sum(x => x.LotQty) } // 单元素数组
                        })
                        .ToList();

                    result = monthStats;
                }
                else {
                    // 按月统计：每个人1-12月的数据（固定12元素数组）
                    var monthStats = dataList
                        .GroupBy(x => x.Inspector)
                        .Select(g => new {
                            name = g.Key,
                            data = Enumerable.Range(1, 12) // 生成1-12月
                                .Select(month => g.Where(x => x.CreateDate.Month == month).Sum(x => x.LotQty))
                                .ToArray() // 转为12元素数组
                        })
                        .ToList();

                    result = monthStats;
                }

                string jsonData = JsonConvert.SerializeObject(result);

                return new ApiResponse
                {
                    Success = true,
                    Message = "数据获取成功",
                    Data = jsonData
                };
            }
            catch (Exception ex)
            {
                return new ApiResponse
                {
                    Success = false,
                    Message = $"数据获取失败：{ex.Message}"
                };
            }
        }

        /// <summary>
        /// 人员检验效率
        /// </summary>
        public async Task<ApiResponse> GetPersonnelEfficiencyDataAsync(INSPECT_PERSONNELDATA input)
        {
            try {
                //INSPECTOR,INSPECT_IQCCREATEDATE,LOT_QTY
                DataTable dt = await GetPersonnelData(input);
                // 转换为可枚举数据并处理空值
                var dataList = dt.AsEnumerable()
                    .Select(row => new {
                        Inspector = row.IsNull("INSPECTOR") ? string.Empty : row.Field<string>("INSPECTOR"),
                        CreateDate = row.IsNull("INSPECT_IQCCREATEDATE") ? DateTime.MinValue : row.Field<DateTime>("INSPECT_IQCCREATEDATE"),
                        LotQty = row.IsNull("LOT_QTY") ? 0 : Convert.ToDecimal(row.Field<object>("LOT_QTY"))
                    })
                    .Where(x => !string.IsNullOrEmpty(x.Inspector) && x.CreateDate != DateTime.MinValue)
                    .ToList();

                object result = null;

                if (input.SumType.ToLower() == "month") {
                    var monthStats = dataList
                        .GroupBy(x => x.Inspector)
                        .Select(g => new {
                            name = g.Key,
                            data = new[] { g.Sum(x => x.LotQty) } // 单元素数组
                        })
                        .ToList();

                    result = monthStats;
                }
                else {
                    // 按月统计：每个人1-12月的数据（固定12元素数组）
                    var monthStats = dataList
                        .GroupBy(x => x.Inspector)
                        .Select(g => new {
                            name = g.Key,
                            data = Enumerable.Range(1, 12) // 生成1-12月
                                .Select(month => g.Where(x => x.CreateDate.Month == month).Sum(x => x.LotQty))
                                .ToArray() // 转为12元素数组
                        })
                        .ToList();

                    result = monthStats;
                }

                string jsonData = JsonConvert.SerializeObject(result);

                return new ApiResponse {
                    Success = true,
                    Message = "数据获取成功",
                    Data = jsonData
                };
            }
            catch (Exception ex) {
                return new ApiResponse {
                    Success = false,
                    Message = $"数据获取失败：{ex.Message}"
                };
            }
        }

        /// <summary>
        /// 人员总检验时长
        /// </summary>
        public async Task<ApiResponse> GetPersonnelDurationDataAsync(INSPECT_PERSONNELDATA input)
        {
            try {
                //INSPECTOR,ITEMKIND,LOT_QTY
                DataTable dt = await GetPersonnelData(input);
                // 定义ITEMKIND对应的检验时长系数（严格匹配名称）
                var itemKindCoefficient = new Dictionary<string, decimal>
                {
                    { "塑胶件", 3 },
                    { "金属件", 3 },
                    { "电子件", 0.5m },
                    { "辅料", 1 },
                    { "包材", 0.5m },
                    { "成品外购件", 2 },
                    { "", 0 }
                };

                // 转换数据并计算每条记录的检验时长
                var dataList = dt.AsEnumerable()
                    .Select(row => new {
                        Inspector = row.IsNull("INSPECTOR") ? string.Empty : row.Field<string>("INSPECTOR"),
                        ItemKind = row.IsNull("ITEMKIND") ? string.Empty : row.Field<string>("ITEMKIND"),
                        LotQty = row.IsNull("LOT_QTY") ? 0 : Convert.ToDecimal(row.Field<object>("LOT_QTY")),
                        CreateDate = row.IsNull("INSPECT_IQCCREATEDATE") ? DateTime.MinValue : row.Field<DateTime>("INSPECT_IQCCREATEDATE"),
                        // 计算单条记录的检验时长：LOT_QTY * 对应系数
                        InspectionTime = row.IsNull("LOT_QTY") ? 0 :
                            Convert.ToDecimal(row.Field<object>("LOT_QTY")) *
                            (itemKindCoefficient.TryGetValue(row.IsNull("ITEMKIND") ? string.Empty : row.Field<string>("ITEMKIND"), out var coeff) ? coeff : 0)
                    })
                    .Where(x => !string.IsNullOrEmpty(x.Inspector) && x.CreateDate != DateTime.MinValue)
                    .ToList();

                object result = null;

                if (input.SumType.ToLower() == "month") {
                    var monthStats = dataList
                       .GroupBy(x => x.Inspector)
                       .Select(g => new {
                           name = g.Key,
                           data = new[] { g.Sum(x => x.InspectionTime) } // 月度总检验时长
                       })
                       .ToList();

                    result = monthStats;
                }
                else {
                    //每个人1-12月的检验时长（固定12元素数组）
                    var monthStats = dataList
                    .GroupBy(x => x.Inspector)
                    .Select(g => new {
                        name = g.Key,
                        data = Enumerable.Range(1, 12)
                            .Select(month => g.Where(x => x.CreateDate.Month == month).Sum(x => x.InspectionTime))
                            .ToArray() // 1-12月检验时长数组
                    })
                    .ToList();

                    result = monthStats;
                }

                string jsonData = JsonConvert.SerializeObject(result);

                return new ApiResponse {
                    Success = true,
                    Message = "数据获取成功",
                    Data = jsonData
                };
            }
            catch (Exception ex) {
                return new ApiResponse {
                    Success = false,
                    Message = $"数据获取失败：{ex.Message}"
                };
            }
        }
        #endregion
    }
}
