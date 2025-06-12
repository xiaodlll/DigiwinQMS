using Mapster;
using MathNet.Numerics.LinearAlgebra.Factorization;
using Meiam.System.Common;
using Meiam.System.Extensions;
using Meiam.System.Extensions.Dto;
using Meiam.System.Hostd.Extensions;
using Meiam.System.Interfaces;
using Meiam.System.Model;
using Meiam.System.Model.Dto;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.IdentityModel.Tokens;
using Newtonsoft.Json;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SqlSugar;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace Meiam.System.Hostd.Controllers.Bisuness
{
    /// <summary>
    /// MS
    /// </summary>
    [Route("api/[controller]/[action]")]
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

    }
}
