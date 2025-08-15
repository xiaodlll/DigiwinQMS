using Meiam.System.Common;
using Meiam.System.Hostd.Controllers.Bisuness;
using Meiam.System.Interfaces;
using Meiam.System.Interfaces.Extensions;
using Meiam.System.Interfaces.IService;
using Meiam.System.Model;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
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
        public async Task<IActionResult> PostLotNotice([FromBody] List<LotNoticeRequest> request)
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

        #region 收料检验结果回传ERP
        /// <summary>
        /// 收料检验结果回传ERP(YNK)
        /// </summary>
        /// <returns></returns>
        [HttpPost("UpdateReceiveInspectResultYNK")]
        public async Task<IActionResult> PostLotNoticeSync()
        {
            List<LotNoticeResultRequest> requests = _ynkService.GetQmsLotNoticeResultRequest();
            string erpApiUrl = AppSettings.Configuration["AppSettings:ERPApiAddress"].TrimEnd('/') + "/api/Qms/UpdateReceiveInspectResult";

            try
            {
                foreach (var request in requests)
                {
                    _logger.LogInformation(@$"请求UpdateReceiveInspectResult: erpApiUrl: {erpApiUrl} request: {JsonConvert.SerializeObject(request)}");
                    string postResult = await HttpHelper.PostJsonAsync(erpApiUrl, request);
                    _logger.LogInformation(@$"回传UpdateReceiveInspectResult: postResult: {postResult}");

                    if (postResult.Contains("false"))
                    {
                        return BadRequest(new ApiResponse
                        {
                            Success = false,
                            Message = $"{postResult}"
                        });
                    }
                    else
                    {
                        _ynkService.CallBackQmsLotNoticeResult(request);
                    }
                }
                return Ok(new ApiResponse
                {
                    Success = true,
                    Message = $"调用成功！"
                });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "调用 ERP 接口异常");

                return StatusCode(500, new ApiResponse
                {
                    Success = false,
                    Message = $"系统异常：{erpApiUrl} :  {ex.ToString()}"
                });
            }
        }
        #endregion
    }
}
