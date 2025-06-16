﻿using Autofac.Core;
using Mapster;
using MathNet.Numerics.LinearAlgebra.Factorization;
using Meiam.System.Common;
using Meiam.System.Extensions;
using Meiam.System.Extensions.Dto;
using Meiam.System.Hostd.Extensions;
using Meiam.System.Interfaces;
using Meiam.System.Interfaces.Extensions;
using Meiam.System.Model;
using Meiam.System.Model.Dto;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.IdentityModel.Tokens;
using Newtonsoft.Json;
using NPOI.HPSF;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SqlSugar;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace Meiam.System.Hostd.Controllers.Bisuness
{
    /// <summary>
    /// MS
    /// </summary>
    [Route("api/erp")]
    [ApiController]
    public class MSController : BaseController
    {
        /// <summary>
        /// 日志管理接口
        /// </summary>
        private readonly ILogger<MSController> _logger;

        /// <summary>
        /// 项目Bom接口
        /// </summary>
        private readonly IMSService _msService;


        public MSController(ILogger<MSController> logger, IMSService msService)
        {
            _logger = logger;
            _msService = msService;
        }


        #region ERP收料通知单
        /// <summary>
        /// ERP收料通知单
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        [HttpPost("lotNotice")]
        public async Task<IActionResult> PostLotNotice([FromBody] LotNoticeRequest request)
        {
            _logger.LogInformation("收到收料通知单请求，单号: {ErpArrivedId}", request?.ERP_ARRIVEDID);
            if (!ModelState.IsValid)
            {
                _logger.LogWarning("请求参数验证失败: {@ModelState}", ModelState);

                return BadRequest(new ApiResponse
                {
                    Success = false,
                    Message = "请求参数验证失败",
                    Data = ModelState
                });
            }

            var result = await _msService.ProcessLotNoticeAsync(request);

            if (result.Success)
            {
                _logger.LogInformation("收料通知单处理成功，单号: {ErpArrivedId}", request.ERP_ARRIVEDID);
                return Ok(result);
            }

            _logger.LogError("收料通知单处理失败，单号: {ErpArrivedId}, 原因: {Message}",
                request.ERP_ARRIVEDID, result.Message);
            return BadRequest(result);
        }
        #endregion

        #region 首检单据
        /// <summary>
        /// 首检单据
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        [HttpPost("workorderSync")]
        public async Task<IActionResult> PostWorkOrderSync([FromBody] WorkOrderSyncRequest request)
        {
            _logger.LogInformation("收到工单同步请求，工单号: {MOID}", request?.MOID);

            if (!ModelState.IsValid)
            {
                _logger.LogWarning("无效的请求参数: {@Errors}", ModelState);
                return BadRequest(new ApiResponse
                {
                    Success = false,
                    Message = "参数验证失败",
                    Data = ModelState
                });
            }

            var result = await _msService.ProcessWorkOrderAsync(request);
            return result.Success ? Ok(result) : BadRequest(result);
        }
        #endregion

        #region 收料检验结果回传ERP
        /// <summary>
        /// 收料检验结果回传ERP
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        [HttpPost("UpdateReceiveInspectResult")]
        public async Task<IActionResult> PostLotNoticeSync([FromBody] LotNoticeResultRequest request)
        {
            string erpApiUrl =  AppSettings.Configuration["AppSettings:ERPApiAddress"].TrimEnd('/') + "/api/Qms/UpdateFirstInspectResult";

            if (!ModelState.IsValid)
            {
                _logger.LogWarning("无效的请求参数: {@Errors}", ModelState);
                return BadRequest(new ApiResponse
                {
                    Success = false,
                    Message = "参数验证失败",
                    Data = ModelState
                });
            }

            try
            {
                string postResult = await HttpHelper.PostJsonAsync(erpApiUrl, request);

                var result = JsonConvert.DeserializeObject<ApiResponse>(postResult);

                if (result?.Success == true)
                {
                    return Ok(result);
                }
                else
                {
                    return BadRequest(result ?? new ApiResponse
                    {
                        Success = false,
                        Message = "更新检验结果失败，返回内容无法解析"
                    });
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "调用 ERP 接口异常");

                return StatusCode(500, new ApiResponse
                {
                    Success = false,
                    Message = $"系统异常：{ex.Message}"
                });
            }
        }
        #endregion

        #region 工单首检检验结果回传MES
        /// <summary>
        /// 工单首检检验结果回传MES
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        [HttpPost("UpdateFirstInspectResult")]
        public async Task<IActionResult> PostWorkOrderResultSync([FromBody] WorkOrderResultRequest request)
        {
            string erpApiUrl = AppSettings.Configuration["AppSettings:ERPApiAddress"].TrimEnd('/') + "/api/Qms/UpdateFirstInspectResult";

            if (!ModelState.IsValid)
            {
                _logger.LogWarning("无效的请求参数: {@Errors}", ModelState);
                return BadRequest(new ApiResponse
                {
                    Success = false,
                    Message = "参数验证失败",
                    Data = ModelState
                });
            }

            try
            {
                string postResult = await HttpHelper.PostJsonAsync(erpApiUrl, request);

                var result = JsonConvert.DeserializeObject<ApiResponse>(postResult);

                if (result?.Success == true)
                {
                    return Ok(result);
                }
                else
                {
                    return BadRequest(result ?? new ApiResponse
                    {
                        Success = false,
                        Message = "更新检验结果失败，返回内容无法解析"
                    });
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "调用 ERP 接口异常");

                return StatusCode(500, new ApiResponse
                {
                    Success = false,
                    Message = $"系统异常：{ex.Message}"
                });
            }
        }
        #endregion



        #region 产品检验结果(入库检)
        /// <summary>
        /// 产品检验结果(入库检)
        /// </summary>
        /// <returns></returns>
        [HttpPost("lotCheckResult")]
        public async Task<IActionResult> PostLotCheckResult([FromBody] LotCheckResultRequest request)
        {
            _logger.LogInformation("收到检验结果查询请求，料号: {ITEMID}", request?.ITEMID);

            if (!ModelState.IsValid)
            {
                _logger.LogWarning("参数验证失败: {@Errors}", ModelState);
                return BadRequest(new CheckResultResponse
                {
                    Success = false,
                    Message = "参数格式错误",
                    Result = "未检验"
                });
            }

            var result = await _msService.ProcessLotCheckResult(request);
            return result.Success ? Ok(result) : BadRequest(result);
        }
        #endregion
    }
}
