﻿using Aspose.Pdf.Operators;
using Autofac.Core;
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
    /// HMD
    /// </summary>
    [Route("api/erp")]
    [ApiController]
    public class HMDController : BaseController
    {
        /// <summary>
        /// 日志管理接口
        /// </summary>
        private readonly ILogger<HMDController> _logger;

        /// <summary>
        /// 项目接口
        /// </summary>
        private readonly IHMDService _hmdService;


        public HMDController(ILogger<HMDController> logger, IHMDService hmdService)
        {
            _logger = logger;
            _hmdService = hmdService;
        }

        #region 恒铭达测量数据上传
        /// <summary>
        /// 恒铭达拉力机数据上传
        /// </summary>
        /// <returns></returns>
        [HttpPost("UploadInpectDev1Data")]
        public async Task<IActionResult> UploadInpectDev1Data([FromBody] InspectDev1Entity input) {
            if (!ModelState.IsValid) {
                _logger.LogWarning("无效的请求参数: {@Errors}", ModelState);

                return BadRequest(new ApiResponse {
                    Success = false,
                    Message = $"参数验证失败，原因：{ModelState}"
                });
            }

            var result = await _hmdService.ProcessHMDInpectDev1DataAsync(input);

            if (result.Success) {
                return Ok(result);
            }
            return BadRequest(result);
        }

        /// <summary>
        /// 获取检验项目信息
        /// </summary>
        /// <returns></returns>
        [HttpPost("GetInspectInfoByDocCodeAsync")]
        public async Task<IActionResult> GetInspectInfoByDocCodeAsync([FromBody] INSPECT_REQCODE input) {
            if (!ModelState.IsValid) {
                _logger.LogWarning("无效的请求参数: {@Errors}", ModelState);

                return BadRequest(new ApiResponse {
                    Success = false,
                    Message = $"参数验证失败，原因：{ModelState}"
                });
            }

            var result = await _hmdService.GetInspectInfoByDocCodeAsync(input);

            if (result.Success) {
                return Ok(result);
            }
            return BadRequest(result);
        }

        /// <summary>
        /// 获取检验项目
        /// </summary>
        /// <returns></returns>
        [HttpPost("GetProgressDataByDocCode")]
        public async Task<IActionResult> GetProgressDataByDocCode([FromBody] INSPECT_REQCODE input) {
            if (!ModelState.IsValid) {
                _logger.LogWarning("无效的请求参数: {@Errors}", ModelState);

                return BadRequest(new ApiResponse {
                    Success = false,
                    Message = $"参数验证失败，原因：{ModelState}"
                });
            }

            var result = await _hmdService.GetProgressDataByDocCodeAsync(input);

            if (result.Success) {
                return Ok(result);
            }
            return BadRequest(result);
        }

        /// <summary>
        /// 获取拉力机规格
        /// </summary>
        /// <returns></returns>
        [HttpPost("GetInspectSpecData")]
        public async Task<IActionResult> GetInspectSpecData([FromBody] INSPECT_SYSM002_REQBYID input) {
            if (!ModelState.IsValid) {
                _logger.LogWarning("无效的请求参数: {@Errors}", ModelState);

                return BadRequest(new ApiResponse {
                    Success = false,
                    Message = $"参数验证失败，原因：{ModelState}"
                });
            }

            var result = await _hmdService.GetInspectSpecDataAsync(input);

            if (result.Success) {
                return Ok(result);
            }
            return BadRequest(result);
        }


        /// <summary>
        /// 获取检验项目信息
        /// </summary>
        /// <returns></returns>
        [HttpPost("GetInspectInfoByConditionAsync")]
        public async Task<IActionResult> GetInspectInfoByConditionAsync([FromBody] INSPECT_CONDITION input) {
            if (!ModelState.IsValid) {
                _logger.LogWarning("无效的请求参数: {@Errors}", ModelState);

                return BadRequest(new ApiResponse {
                    Success = false,
                    Message = $"参数验证失败，原因：{ModelState}"
                });
            }

            var result = await _hmdService.GetInspectInfoByConditionAsync(input);

            if (result.Success) {
                return Ok(result);
            }
            return BadRequest(result);
        }

        /// <summary>
        /// 恒铭达二次元数据上传
        /// </summary>
        /// <returns></returns>
        [HttpPost("UploadInpectProcessData")]
        public async Task<IActionResult> UploadInpectProcessData([FromBody] List<INSPECT_PROGRESSDto> input) {
            if (!ModelState.IsValid) {
                _logger.LogWarning("无效的请求参数: {@Errors}", ModelState);

                return BadRequest(new ApiResponse {
                    Success = false,
                    Message = $"参数验证失败，原因：{ModelState}"
                });
            }

            var result = await _hmdService.ProcessHMDInpectProcessDataAsync(input);

            if (result.Success) {
                return Ok(result);
            }
            return BadRequest(result);
        }


        /// <summary>
        /// 获取ROSH检验项目信息
        /// </summary>
        /// <returns></returns>
        [HttpPost("GetInspectInfoByRoshConditionAsync")]
        public async Task<IActionResult> GetInspectInfoByRoshConditionAsync([FromBody] INSPECT_CONDITION input) {
            if (!ModelState.IsValid) {
                _logger.LogWarning("无效的请求参数: {@Errors}", ModelState);

                return BadRequest(new ApiResponse {
                    Success = false,
                    Message = $"参数验证失败，原因：{ModelState}"
                });
            }

            var result = await _hmdService.GetInspectInfoByRoshConditionAsync(input);

            if (result.Success) {
                return Ok(result);
            }
            return BadRequest(result);
        }
        
        /// <summary>
        /// 恒铭达ROSH数据上传
        /// </summary>
        /// <returns></returns>
        [HttpPost("UploadInpectProcessRoshData")]
        public async Task<IActionResult> UploadInpectProcessRoshData([FromBody] HMDRoshDataDto input) {
            if (!ModelState.IsValid) {
                _logger.LogWarning("无效的请求参数: {@Errors}", ModelState);

                return BadRequest(new ApiResponse {
                    Success = false,
                    Message = $"参数验证失败，原因：{ModelState}"
                });
            }

            var result = await _hmdService.ProcessHMDInpectProcessRoshDataAsync(input);

            if (result.Success) {
                return Ok(result);
            }
            return BadRequest(result);
        }

        #endregion

        #region 检验结果回传ERP
        /// <summary>
        /// 收料检验结果回传ERP
        /// </summary>
        /// <returns></returns>
        [HttpPost("UpdateReceiveInspectResultHMD")]
        public async Task<IActionResult> PostLotNoticeSync() {
            try {
                await _hmdService.UpdateReceiveInspectResult();
                return Ok(new ApiResponse {
                    Success = true,
                    Message = $"调用成功！"
                });
            }
            catch (Exception ex) {
                _logger.LogError(ex, "调用 ERP 接口异常");

                return StatusCode(500, new ApiResponse {
                    Success = false,
                    Message = $"调用 ERP 接口异常：{ex.Message}"
                });
            }
        }
        
        /// <summary>
        /// FQC检验结果回传ERP
        /// </summary>
        /// <returns></returns>
        [HttpPost("UpdateFQCResultHMD")]
        public async Task<IActionResult> UpdateFQCResultHMD() {
            try {
                await _hmdService.UpdateFqcResult();
                return Ok(new ApiResponse {
                    Success = true,
                    Message = $"调用成功！"
                });
            }
            catch (Exception ex) {
                _logger.LogError(ex, "调用 ERP 接口异常");

                return StatusCode(500, new ApiResponse {
                    Success = false,
                    Message = $"调用 ERP 接口异常：{ex.Message}"
                });
            }
        }
        #endregion


    }
}
