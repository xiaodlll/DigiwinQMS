using Aspose.Pdf.Operators;
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
        /// 恒铭达测量数据上传
        /// </summary>
        /// <returns></returns>
        [HttpPost("ReceiveInspectData")]
        public async Task<IActionResult> ReceiveInspectData([FromBody] HMDInputDto input) {
            if (!ModelState.IsValid) {
                _logger.LogWarning("无效的请求参数: {@Errors}", ModelState);

                return BadRequest(new ApiResponse {
                    Success = false,
                    Message = $"参数验证失败，原因：{ModelState}"
                });
            }

            var result = await _hmdService.ProcessHMDInspectDataAsync(input);

            if (result.Success) {
                return Ok(result);
            }
            return BadRequest(result);
        }
        #endregion
    }
}
