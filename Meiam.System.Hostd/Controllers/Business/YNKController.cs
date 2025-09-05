﻿using Meiam.System.Common;
using Meiam.System.Hostd.Controllers.Bisuness;
using Meiam.System.Interfaces;
using Meiam.System.Interfaces.Extensions;
using Meiam.System.Interfaces.IService;
using Meiam.System.Interfaces.Service;
using Meiam.System.Model;
using Meiam.System.Model.Dto;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using SqlSugar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Meiam.System.Hostd.Controllers.Business
{
    /// <summary>
    /// YNK
    /// </summary>
    [Route("api/erp")]
    [ApiController]
    public class YNKController : BaseController
    {
        /// <summary>
        /// 日志管理接口
        /// </summary>
        private readonly ILogger<YNKController> _logger;

        /// <summary>
        /// 项目Bom接口
        /// </summary>
        private readonly IYNKService _ynkService;


        public YNKController(ILogger<YNKController> logger, IYNKService ynkService)
        {
            _logger = logger;
            _ynkService = ynkService;
        }

        #region ERP收料通知单
        /// <summary>
        /// ERP收料通知单(YNK)
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        [HttpPost("lotNoticeYNK")]
        public async Task<IActionResult> PostLotNotice([FromBody] List<LotNoticeRequestYNK> request)
        {
            if (!ModelState.IsValid)
            {
                _logger.LogWarning("请求参数验证失败: {@ModelState}", ModelState);

                return BadRequest(new ApiResponse
                {
                    Success = false,
                    Message = $"请求参数验证失败，原因：{ModelState}"
                });
            }

            var result = await _ynkService.ProcessLotNoticeAsync(request);

            if (result.Success)
            {
                return Ok(result);
            }

            _logger.LogError("收料通知单处理失败， 原因: {Message}",
                result.Message);
            return BadRequest(result);
        }
        #endregion

        #region 金蝶云登录接口
        /// <summary>
        /// 金蝶云登录接口
        /// </summary>
        /// <returns>登录结果包含KDSVCSessionId</returns>
        [HttpPost("LoginYNK")]
        [ProducesResponseType(typeof(ApiResponseYNK<ERPLoginResponseYNK>), 200)]
        [ProducesResponseType(typeof(ApiErrorResponseYNK), 400)]
        public async Task<IActionResult> Login()
        {
            var result = await _ynkService.LoginAsync();

            if (result.IsSuccess)
            {
                return Ok(new ApiResponseYNK<ERPLoginResponseYNK>
                {
                    Success = true,
                    Data = result,
                    Message = "登录成功"
                });
            }
            else
            {
                return BadRequest(new ApiErrorResponseYNK
                {
                    Success = false,
                    ErrorCode = "LOGIN_FAILED",
                    ErrorMessage = result.ErrorMessage,
                    StatusCode = result.StatusCode
                });
            }
        }

        /// <summary>
        /// 自定义参数的金蝶云登录接口
        /// </summary>
        /// <param name="request">登录参数</param>
        /// <returns>登录结果包含KDSVCSessionId</returns>
        [HttpPost("LoginYNK/custom")]
        [ProducesResponseType(typeof(ApiResponseYNK<ERPLoginResponseYNK>), 200)]
        [ProducesResponseType(typeof(ApiErrorResponseYNK), 400)]
        public async Task<IActionResult> LoginWithCustom([FromBody] ERPLoginRequestYNK request)
        {
            if (string.IsNullOrEmpty(request.Username) ||
                string.IsNullOrEmpty(request.Password) ||
                string.IsNullOrEmpty(request.AcctID))
            {
                return BadRequest(new ApiErrorResponseYNK
                {
                    Success = false,
                    ErrorCode = "INVALID_PARAMS",
                    ErrorMessage = "用户名、密码和账套ID不能为空",
                    StatusCode = 400
                });
            }

            var result = await _ynkService.LoginAsync(
                request.Username,
                request.Password,
                request.AcctID,
                request.Lcid
            );

            if (result.IsSuccess)
            {
                return Ok(new ApiResponseYNK<ERPLoginResponseYNK>
                {
                    Success = true,
                    Data = result,
                    Message = "登录成功"
                });
            }
            else
            {
                return BadRequest(new ApiErrorResponseYNK
                {
                    Success = false,
                    ErrorCode = "LOGIN_FAILED",
                    ErrorMessage = result.ErrorMessage,
                    StatusCode = result.StatusCode
                });
            }
        }

        #endregion

        #region 收料检验结果回传ERP
        /// <summary>
        /// 收料检验结果回传ERP(YNK)
        /// </summary>
        /// <returns></returns>
        [HttpPost("UpdateReceiveInspectResultYNK")]
        public async Task<IActionResult> PostLotNoticeSync()
        {
            List<LotNoticeResultRequestYNK> requests = _ynkService.GetQmsLotNoticeResultRequest();

            string erpApiUrl = AppSettings.Configuration["ERP:BaseUrl"] + AppSettings.Configuration["ERP:LoginUrl"];

            try
            {
                // 1. 首先获取KDSVCSessionId
                var loginResult = await _ynkService.LoginAsync();

                if (!loginResult.IsSuccess || string.IsNullOrEmpty(loginResult.KDSVCSessionId))
                {
                    _logger.LogError($"获取KDSVCSessionId失败: {loginResult.ErrorMessage}");
                    return BadRequest(new ApiResponse
                    {
                        Success = false,
                        Message = $"登录ERP系统失败: {loginResult.ErrorMessage}"
                    });
                }

                string sessionId = loginResult.KDSVCSessionId;
                _logger.LogInformation($"成功获取KDSVCSessionId: {sessionId}");

                //按FID将分组
                var groupedByFid = requests.GroupBy(r => r.FID)
                                    .Where(g => !string.IsNullOrEmpty(g.Key)) // 确保FID非空
                                    .ToList();

                _logger.LogInformation($"按FID分组完成，共 {groupedByFid.Count} 个单据");

                int successCount = 0;
                int failCount = 0;

                // 分组后的数据一单一单的传
                foreach (var group in groupedByFid)
                {
                    var fid = group.Key;
                    var entries = group.ToList();

                    try
                    {
                        // 封装成金蝶ERP需要的格式
                        var erpRequestData = new
                        {
                            formid = "PUR_ReceiveBill", 
                            data = new
                            {
                                NeedUpDateFields = new[] { "FDetailEntity", "FReceiveQty", "FRefuseQty", "FCheckQty" },
                                IsDeleteEntry = "false",
                                IsVerifyBaseDataField = "true",
                                IsAutoAdjustField = "true",
                                Model = new
                                {
                                    FID = fid, // 传进来的FID
                                    FDetailEntity = entries.Select(entry => new
                                    {
                                        FEntryID = entry.FEntryID, // 传进来的FEntryID
                                        FReceiveQty = entry.FReceiveQty, // 传进来的FReceiveQty
                                        FRefuseQty = entry.FRefuseQty, // 传进来的FRefuseQty
                                        FCheckQty = entry.FReceiveQty + entry.FRefuseQty // FReceiveQty和FRefuseQty相加
                                    }).ToList()
                                }
                            }
                        };

                        _logger.LogInformation(@$"请求金蝶ERP接口: FID: {fid}, 包含 {entries.Count} 个明细行");

                        string jsonRequest = JsonConvert.SerializeObject(erpRequestData);
                        _logger.LogInformation(@$"请求URL: {erpApiUrl}");
                        _logger.LogInformation(@$"请求数据: {jsonRequest}");

                        // 使用带有SessionId的HTTP请求
                        string postResult = await HttpHelper.PostJsonWithSessionAsync(
                            erpApiUrl,
                            jsonRequest,
                            sessionId
                        );

                        _logger.LogInformation(@$"金蝶ERP接口响应 - FID: {fid}, 结果: {postResult}");

                        // 解析响应结果
                        if (postResult.Contains("false"))
                        {
                            failCount++;
                            _logger.LogError($"单据 {fid} 回传失败: {postResult}");
                        }
                        else
                        {
                            // 回传成功，更新所有相关明细行的状态
                            foreach (var entry in entries)
                            {
                                _ynkService.CallBackQmsLotNoticeResult(entry);
                            }
                            successCount++;
                            _logger.LogInformation($"单据 {fid} 回传成功");
                        }
                    }
                    catch (Exception ex)
                    {
                        failCount++;
                        _logger.LogError(ex, $"处理单据 {fid} 时发生异常:" + ex.ToString());
                        // 继续处理其他单据
                        continue;
                    }
                }

                return Ok(new ApiResponse
                {
                    Success = true,
                    Message = $"处理完成！成功: {successCount} 个单据, 失败: {failCount} 个单据, 总计: {groupedByFid.Count} 个单据"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "调用 ERP 接口异常");

                return StatusCode(500, new ApiResponse
                {
                    Success = false,
                    Message = $"系统异常：{erpApiUrl} : {ex.Message}"
                });
            }
        }
        #endregion

        #region 工具API
        /// <summary>
        /// 获取检验项目信息
        /// </summary>
        /// <returns></returns>
        [HttpPost("GetAOIInspectInfoByDocCode")]
        public async Task<IActionResult> GetAOIInspectInfoByDocCode([FromBody] INSPECT_REQCODE input) {
            if (!ModelState.IsValid) {
                _logger.LogWarning("无效的请求参数: {@Errors}", ModelState);

                return BadRequest(new ApiResponse {
                    Success = false,
                    Message = $"参数验证失败，原因：{ModelState}"
                });
            }

            var result = await _ynkService.GetAOIInspectInfoByDocCodeAsync(input);

            if (result.Success) {
                return Ok(result);
            }
            return BadRequest(result);
        }


        /// <summary>
        /// 获取检验项目
        /// </summary>
        /// <returns></returns>
        [HttpPost("GetAOIProgressDataByDocCode")]
        public async Task<IActionResult> GetAOIProgressDataByDocCode([FromBody] INSPECT_REQCODE input) {
            if (!ModelState.IsValid) {
                _logger.LogWarning("无效的请求参数: {@Errors}", ModelState);

                return BadRequest(new ApiResponse {
                    Success = false,
                    Message = $"参数验证失败，原因：{ModelState}"
                });
            }

            var result = await _ynkService.GetAOIProgressDataByDocCodeAsync(input);

            if (result.Success) {
                return Ok(result);
            }
            return BadRequest(result);
        }

        /// <summary>
        /// AOI数据上传
        /// </summary>
        /// <returns></returns>
        [HttpPost("UploadAOIData")]
        public async Task<IActionResult> UploadAOIData([FromBody] List<InspectAoi> input) {
            if (!ModelState.IsValid) {
                _logger.LogWarning("无效的请求参数: {@Errors}", ModelState);

                return BadRequest(new ApiResponse {
                    Success = false,
                    Message = $"参数验证失败，原因：{ModelState}"
                });
            }

            var result = await _ynkService.ProcessUploadAOIDataAsync(input);

            if (result.Success) {
                return Ok(result);
            }
            return BadRequest(result);
        }

        /// <summary>
        /// AOI图片数据上传
        /// </summary>
        /// <returns></returns>
        [HttpPost("UploadAOIImageData")]
        public async Task<IActionResult> UploadAOIImageData([FromBody] List<InspectImageAoi> input) {
            if (!ModelState.IsValid) {
                _logger.LogWarning("无效的请求参数: {@Errors}", ModelState);

                return BadRequest(new ApiResponse {
                    Success = false,
                    Message = $"参数验证失败，原因：{ModelState}"
                });
            }

            var result = await _ynkService.ProcessUploadAOIImageDataAsync(input);

            if (result.Success) {
                return Ok(result);
            }
            return BadRequest(result);
        }

        /// <summary>
        /// 二次元数据上传
        /// </summary>
        /// <returns></returns>
        [HttpPost("UploadYNKInpectProcessData")]
        public async Task<IActionResult> UploadYNKInpectProcessData([FromBody] List<INSPECT_PROGRESSDto> input) {
            if (!ModelState.IsValid) {
                _logger.LogWarning("无效的请求参数: {@Errors}", ModelState);

                return BadRequest(new ApiResponse {
                    Success = false,
                    Message = $"参数验证失败，原因：{ModelState}"
                });
            }

            var result = await _ynkService.ProcessYNKInpectProcessDataAsync(input);

            if (result.Success) {
                return Ok(result);
            }
            return BadRequest(result);
        }
        #endregion
    }
}
