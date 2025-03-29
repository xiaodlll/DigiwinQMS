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
            var defaultData = _iqcService.GetWhere(m => m.INSPECT_DEV1ID == parm.INSPECT_DEV1ID).FirstOrDefault();
            if (defaultData == null) {
                return toResponse(StatusCodeType.Error, $"到不到{parm.INSPECT_DEV1ID}对应的数据！");
            }
            if(_iqcService.ExistScanDoc("拉力机检测图", parm.INSPECT_IQCCODE)) {
                return toResponse(StatusCodeType.Error, $"拉力机检测图{parm.INSPECT_IQCCODE}在数据库已存在！");
            }

            string itemID = defaultData.ITEMID;
            string lotID = defaultData.LOTID;
            string inspect_Date = defaultData.INSPECT_DATE.Value.ToString("yyyyMMddHHmmss");

            var allDBDataCount = _iqcService.GetCount(m => m.INSPECT_DEV1ID == parm.INSPECT_DEV1ID);
            var predicate = Expressionable.Create<INSPECT_TENSILE>();

            List<INSPECT_TENSILE_D> listToSave = new List<INSPECT_TENSILE_D>();
            if (parm.YSAMPLE > allDBDataCount) {
                predicate = predicate.And(m => m.INSPECT_DEV1ID != parm.INSPECT_DEV1ID && m.ITEMID == itemID);

                var addDBData = _iqcService.GetWhere(m => m.INSPECT_DEV1ID == parm.INSPECT_DEV1ID);
                foreach (var item in addDBData) {
                    listToSave.Add(GetDetailByInspect(item));
                }
                var addDBDataEx = _iqcService.GetRandomData(predicate.ToExpression(), parm.YSAMPLE - allDBDataCount);
                if (addDBDataEx.Count() > 0) {
                    foreach (var item in addDBDataEx) {
                        var itemEx = GetDetailByInspect(item);
                        itemEx.Flag = true;
                        listToSave.Add(itemEx);
                    }
                }
            }
            else {
                predicate = predicate.And(m => m.INSPECT_DEV1ID == parm.INSPECT_DEV1ID);
                var addDBData = _iqcService.GetRandomData(predicate.ToExpression(), parm.YSAMPLE);
                foreach (var item in addDBData) {
                    listToSave.Add(GetDetailByInspect(item));
                }
            }
            _iqcService.SaveToInspectDetail(listToSave);

            byte[] fileContents = _iqcService.GetInspectImage(listToSave);

            //返回文件流
            var fileName = $"{itemID}_{lotID}{inspect_Date}.jpg";
            string filePath = Path.Combine(AppSettings.Configuration["AppSettings:FileServerPath"], @$"TENSILE\{itemID}\{fileName}");
            
            //保存到SCANDOC
            _iqcService.SaveToScanDoc("拉力机检测图", fileContents, filePath, parm.INSPECT_IQCCODE);

            return File(fileContents, "image/jpeg", fileName);
        }

        private INSPECT_TENSILE_D GetDetailByInspect(INSPECT_TENSILE item) {
            var newItem = new INSPECT_TENSILE_D();
            newItem.INSPECT_TENSILE_DID = Guid.NewGuid().ToString();
            // 使用反射复制属性值
            var sourceProperties = item.GetType().GetProperties();
            foreach (var property in sourceProperties) {
                var targetProperty = newItem.GetType().GetProperty(property.Name);
                if (targetProperty != null && targetProperty.CanWrite) {
                    targetProperty.SetValue(newItem, property.GetValue(item));
                }
            }
            return newItem;
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



            //byte[] fileContents = _iqcService.getcpk(listToSave);


            //byte[] fileContents = _iqcService.GetCPKReport(parm);
            byte[] fileContents = null;
            // 3. 返回文件流
            var fileName = $"零件清单_{DateTime.Now:yyyyMMdd}.jpg";
            return File(fileContents, "image/jpeg", fileName);
        }
    }
}
