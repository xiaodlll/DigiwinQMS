using Meiam.System.Hostd.Extensions;
using Meiam.System.Interfaces;
using Meiam.System.Model;
using Meiam.System.Model.Dto;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using SqlSugar;

namespace Meiam.System.Hostd.Controllers.System
{
    /// <summary>
    /// 日志接口
    /// </summary>
    [Route("api/[controller]/[action]")]
    [ApiController]
    public class LogsController : BaseController
    {
        /// <summary>
        /// 日志管理接口
        /// </summary>
        private readonly ILogger<LogsController> _logger;

        /// <summary>
        /// 日志接口
        /// </summary>
        private readonly ISysLogsService _logsService;

        public LogsController(ILogger<LogsController> logger, ISysLogsService logsService)
        {
            _logger = logger;
            _logsService = logsService;
        }


        /// <summary>
        /// 查询日志列表
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public IActionResult Query([FromBody] LogsQueryDto parm)
        {
            //开始拼装查询条件
            var predicate = Expressionable.Create<Sys_Logs>();

            predicate = predicate.And(m => m.CreateTime >= parm.BeginDate && m.CreateTime < parm.EndDate.AddDays(1));
            predicate = predicate.AndIF(!string.IsNullOrEmpty(parm.Level), m => m.Level == parm.Level);
            predicate = predicate.AndIF(!string.IsNullOrEmpty(parm.QueryText), m => m.Logger.Contains(parm.QueryText) || m.Message.Contains(parm.QueryText) || m.Url.Contains(parm.QueryText) || m.IPAddress.Contains(parm.QueryText));

            var response = _logsService.GetPages(predicate.ToExpression(), parm);

            return toResponse(response);
        }


    }
}
