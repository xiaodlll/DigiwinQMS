using Mapster;
using Meiam.System.Common;
using Meiam.System.Extensions;
using Meiam.System.Extensions.Dto;
using Meiam.System.Hostd.Extensions;
using Meiam.System.Interfaces;
using Meiam.System.Model;
using Meiam.System.Model.Dto;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using SqlSugar;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace Meiam.System.Hostd.Controllers.Bisuness {
    /// <summary>
    /// IQC
    /// </summary>
    [Route("api/[controller]/[action]")]
    [ApiController]
    public class IQCController : BaseController {
        /// <summary>
        /// 日志管理接口
        /// </summary>
        private readonly ILogger<IQCController> _logger;

        /// <summary>
        /// 项目Bom接口
        /// </summary>
        private readonly IIQCService _iqcService;


        public IQCController(ILogger<IQCController> logger, IIQCService iqcService) {
            _logger = logger;
            _iqcService = iqcService;
        }

        /// <summary>
        /// 拉力机检测报告
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public IActionResult InspectReport([FromBody] InspectInputDto parm) {
            if (string.IsNullOrEmpty(parm.INSPECT_DEV1ID)) {
                return toResponse(StatusCodeType.Error, $"INSPECT_DEV1ID不能为空！");
            }
            if (string.IsNullOrEmpty(parm.UserName))
            {
                return toResponse(StatusCodeType.Error, $"UserName不能为空！");
            }

            byte[] fileContents = _iqcService.GetInspectReport(parm);

            return File(fileContents, "image/jpeg", "拉力机检测报告.jpg");
        }

        /// <summary>
        /// CPK数据报告
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public IActionResult CPKReport([FromBody] CPKInputDto parm) {

            if (string.IsNullOrEmpty(parm.INSPECT_DEV2ID))
            {
                return toResponse(StatusCodeType.Error, $"INSPECT_DEV2ID不能为空！");
            }

            if (string.IsNullOrEmpty(parm.UserName))
            {
                return toResponse(StatusCodeType.Error, $"UserName不能为空！");
            }

            byte[] fileContents = _iqcService.GetCPKfile(parm.INSPECT_DEV2ID, parm.UserName);

            //byte[] fileContents = _iqcService.getcpk(listToSave);


            //byte[] fileContents = _iqcService.GetCPKReport(parm);
            //byte[] fileContents = null;
            // 3. 返回文件流
            var fileName = $"零件清单_{DateTime.Now:yyyyMMdd}.jpg";
            return File(fileContents, "image/jpeg", fileName);
        }
    }
}
