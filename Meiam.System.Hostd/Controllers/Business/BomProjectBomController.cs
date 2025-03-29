using Mapster;
using Meiam.System.Extensions;
using Meiam.System.Extensions.Dto;
using Meiam.System.Hostd.Authorization;
using Meiam.System.Hostd.Extensions;
using Meiam.System.Interfaces;
using Meiam.System.Model;
using Meiam.System.Model.Dto;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using SqlSugar;
using System;

namespace Meiam.System.Hostd.Controllers.Bisuness
{
    /// <summary>
    /// 项目Bom
    /// </summary>
    [Route("api/[controller]/[action]")]
    [ApiController]
    public class BomProjectBomController : BaseController
    {
        /// <summary>
        /// 日志管理接口
        /// </summary>
        private readonly ILogger<BomProjectBomController> _logger;
        /// <summary>
        /// 会话管理接口
        /// </summary>
        private readonly TokenManager _tokenManager;

        /// <summary>
        /// 项目Bom接口
        /// </summary>
        private readonly IBomProjectBomService _bomProjectBomService;


        public BomProjectBomController(ILogger<BomProjectBomController> logger, TokenManager tokenManager, IBomProjectBomService bomProjectBomService)
        {
            _logger = logger;
            _tokenManager = tokenManager;
            _bomProjectBomService = bomProjectBomService;
        }

        /// <summary>
        /// 查询项目Bom列表
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        //[Authorization(Power = "PRIV_PROJECTBOM_VIEW")]
        public IActionResult Query([FromBody] ProjectBomQueryDto parm)
        {
            //开始拼装查询条件
            var predicate = Expressionable.Create<Bom_ProjectBom>();
            predicate = predicate.AndIF(!string.IsNullOrEmpty(parm.ProjectID), m => m.ProjectID == parm.ProjectID);
            predicate = predicate.AndIF(!string.IsNullOrEmpty(parm.Version), m => m.Version ==  parm.Version);
            var response = _bomProjectBomService.GetPages(predicate.ToExpression(), parm);
            return toResponse(response);
        }


        /// <summary>
        /// 查询项目信息
        /// </summary>
        /// <param name="id">ID</param>
        /// <returns></returns>
        [HttpGet]
        //[Authorization(Power = "PRIV_PROJECTBOM_VIEW")]
        public IActionResult Get(string id)
        {
            if (!string.IsNullOrEmpty(id))
            {
                return toResponse(_bomProjectBomService.GetId(id));
            }
            return toResponse(string.Empty);
        }


        /// <summary>
        /// 添加项目Bom
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        //[Authorization(Power = "PRIV_PROJECTBOM_CREATE")]
        public IActionResult Create([FromBody] ProjectBomCreateDto parm)
        {
            if (_bomProjectBomService.Any(m => m.ProjectID == parm.ProjectID && m.Version == parm.Version))
            {
                return toResponse(StatusCodeType.Error, $"添加 {parm.ProjectID} {parm.Version}失败，该项目Bom已存在，不能重复！");
            }

            //从 Dto 映射到 实体
            var entity = parm.Adapt<Bom_ProjectBom>().ToCreate(_tokenManager.GetSessionInfo());

            return toResponse(_bomProjectBomService.Add(entity));
        }

        /// <summary>
        /// 更新项目Bom
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        //[Authorization(Power = "PRIV_PROJECTBOM_UPDATE")]
        public IActionResult Update([FromBody] ProjectBomUpdateDto parm)
        {
            if (_bomProjectBomService.Any(m => m.ProjectID == parm.ProjectID && m.Version == parm.Version && m.ID != parm.ID))
            {
                return toResponse(StatusCodeType.Error, $"添加 {parm.ProjectID} {parm.Version}失败，该项目Bom已存在，不能重复！");
            }

            var userSession = _tokenManager.GetSessionInfo();

            return toResponse(_bomProjectBomService.Update(m => m.ID == parm.ID, m => new Bom_ProjectBom()
            {
                ProjectID = parm.ProjectID,
                Version = parm.Version,
                PriceFinishTime = parm.PriceFinishTime,
                PriceFinishPercent = parm.PriceFinishPercent,
                Remark = parm.Remark,
                UpdateID = userSession.UserID,
                UpdateTime = DateTime.Now
            }));
        }

        /// <summary>
        /// 删除项目信息
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        //[Authorization(Power = "PRIV_PROJECTBOM_DELETE")]
        public IActionResult Delete(string id)
        {
            if (string.IsNullOrEmpty(id))
            {
                return toResponse(StatusCodeType.Error, "删除项目Bom Id 不能为空");
            }

            var response = _bomProjectBomService.Delete(id);

            return toResponse(response);
        }

    }
}
