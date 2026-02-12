using DocumentFormat.OpenXml.Spreadsheet;
using iText.IO.Image;
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
using System.IO;
using System.Data;
using System.Linq;
using OfficeOpenXml.Drawing;
using OfficeOpenXml;
using System.Xml;

namespace Meiam.System.Hostd.Controllers.Bisuness {
    /// <summary>
    /// IQC
    /// </summary>
    [Route("api/[controller]/[action]")]
    [ApiController]
    public class IQCController : BaseController
    {
        /// <summary>
        /// 日志管理接口
        /// </summary>
        private readonly ILogger<IQCController> _logger;

        /// <summary>
        /// 项目Bom接口
        /// </summary>
        private readonly IIQCService _iqcService;


        public IQCController(ILogger<IQCController> logger, IIQCService iqcService)
        {
            _logger = logger;
            _iqcService = iqcService;
        }

        /// <summary>
        /// 拉力机检测报告
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public IActionResult InspectReport([FromBody] InspectInputDto parm)
        {
            if (string.IsNullOrEmpty(parm.INSPECT_DEV1ID))
            {
                return toResponse(StatusCodeType.Error, $"INSPECT_DEV1ID不能为空！");
            }
            if (string.IsNullOrEmpty(parm.UserName))
            {
                return toResponse(StatusCodeType.Error, $"UserName不能为空！");
            }

            try
            {
                _iqcService.GetInspectReport(parm);
            }
            catch (Exception ex)
            {
                return toResponse(StatusCodeType.Error, ex.ToString());
            }
            return toResponse(StatusCodeType.Success, "拉力机检测报告生成成功！");

            //byte[] fileContents = _iqcService.GetInspectReport(parm);

            //return File(fileContents, "image/jpeg", "拉力机检测报告.jpg");
        }

        [HttpPost]
        public IActionResult DEV1_UNION([FromBody] InspectInputDto parm)
        {
            
            try
            {
                _iqcService.GET_DEV1_UNION("");
            }
            catch (Exception ex)
            {
                return toResponse(StatusCodeType.Error, ex.ToString());
            }
            return toResponse(StatusCodeType.Success,  "关联成功！");

        }


        /// <summary>
        /// 产生检验单随机值
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public IActionResult INSPECT_VIEW_RANK([FromBody] InspectInputByCodeDto parm)
        {

            try
            {

                _iqcService.INSPECT_VIEW_RANK(parm.DOC_CODE);

            }
            catch (Exception ex)
            {
                return toResponse(StatusCodeType.Error, parm.DOC_CODE + "∫" + ex.ToString());
            }
            return toResponse(StatusCodeType.Success, parm.DOC_CODE + "∫" + "产生成功！");

        }
        /// <summary>
        /// 批量拉力机检测报告
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public IActionResult InspectBatchReport([FromBody] InspectInputByCodeDto parm)
        {
            if (string.IsNullOrEmpty(parm.DOC_CODE))
            {
                return toResponse(StatusCodeType.Error, $"DOC_CODE不能为空！");
            }
            if (string.IsNullOrEmpty(parm.INSPECT_DEV))
            {
                return toResponse(StatusCodeType.Error, $"INSPECT_DEV不能为空！");
            }
            if (string.IsNullOrEmpty(parm.UserName))
            {
                return toResponse(StatusCodeType.Error, $"UserName不能为空！");
            }

            _iqcService.GetInspectBatchReport(parm);

            return toResponse(StatusCodeType.Success, "拉力机检测报告生成成功！");

            //byte[] fileContents = _iqcService.GetInspectReport(parm);

            //return File(fileContents, "image/jpeg", "拉力机检测报告.jpg");
        }

        /// <summary>
        /// CPK数据报告
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public IActionResult CPKReport([FromBody] CPKInputDto parm)
        {

            if (string.IsNullOrEmpty(parm.INSPECT_DEV2ID))
            {
                return toResponse(StatusCodeType.Error, $"INSPECT_DEV2ID不能为空！");
            }

            if (string.IsNullOrEmpty(parm.UserName))
            {
                return toResponse(StatusCodeType.Error, $"UserName不能为空！");
            }
            try
            {
                _iqcService.GetCPKfile(parm.INSPECT_DEV2ID, parm.UserName);
            }
            catch(Exception ex)
            {
                return toResponse(StatusCodeType.Error, ex.ToString());
            }
            return toResponse(StatusCodeType.Success, "CPK数据报告生成成功！");

            //byte[] fileContents = _iqcService.getcpk(listToSave);


            //byte[] fileContents = _iqcService.GetCPKReport(parm);
            //byte[] fileContents = null;
            //// 3. 返回文件流
            //var fileName = $"零件清单_{DateTime.Now:yyyyMMdd}.jpg";
            //return File(fileContents, "image/jpeg", fileName);
        }

        /// <summary>
        /// 批量CPK数据报告
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public IActionResult CPKBatchReport([FromBody] CPKInputByCodeDto parm)
        {
            if (string.IsNullOrEmpty(parm.DOC_CODE))
            {
                return toResponse(StatusCodeType.Error, $"DOC_CODE不能为空！");
            }
            if (string.IsNullOrEmpty(parm.INSPECT_DEV))
            {
                return toResponse(StatusCodeType.Error, $"INSPECT_DEV不能为空！");
            }
            if (string.IsNullOrEmpty(parm.UserName))
            {
                return toResponse(StatusCodeType.Error, $"UserName不能为空！");
            }

            _iqcService.GetBatchCPKfile(parm);

            return toResponse(StatusCodeType.Success, "拉力机检测报告生成成功！");

            //byte[] fileContents = _iqcService.GetInspectReport(parm);

            //return File(fileContents, "image/jpeg", "拉力机检测报告.jpg");
        }

        /// <summary>
        /// CPK数据报告测试Excel
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public IActionResult TestCPKExcel()
        {
            bool.TryParse(AppSettings.Configuration["AppSettings:IQCReplaceFormula"], out bool isReplaceFormula);
            string filePath = @"C:\Users\Administrator\Desktop\Temp\222.xlsx";
            string filePath2 = @"C:\Users\Administrator\Desktop\Temp\22.doc";
            string[] sourceExcelPaths = new string[] {
                @"C:/Users/Administrator/Desktop/Temp/2.5.xlsx"};
            List<ExcelAttechFile> excelAttechFilesList = new List<ExcelAttechFile>();
            excelAttechFilesList.Add(new ExcelAttechFile() { FileName = null, FilePath = sourceExcelPaths[0] });

            try {
                using (ExcelHelper excelHelper = new ExcelHelper(filePath)) {

                    //excelHelper.CopyColumnsCells("CPK", "C21:D30", 5);
                    //excelHelper.CopySheet(sourceExcelPaths, "ROS", "A1");
                    //excelHelper.CopyRows("出货报告", new int[] { 39 }, 40);
                    //excelHelper.CopyRows("原材料COC", "A6:O6", 3);
                    //string pdfPath1 = excelHelper.ConvertExcelToPdf(filePath);
                    //string pdfPath2 = excelHelper.ConvertWordToPdf(filePath2);
                    excelHelper.AddAttachsToCell("FTIR", "B7", excelAttechFilesList.ToArray(), false, true, false);
                    //excelHelper.InsertRows("出货报告", 3, 40);
                    #region 替换公式
                    //if (isReplaceFormula) {
                        excelHelper.ReplaceFormula();
                    //}
                    #endregion
                }

                return toResponse(StatusCodeType.Success, "测试成功!");
            }
            catch (Exception ex) {
                return toResponse(StatusCodeType.Error, ex.ToString());
            }
        }
        /// <summary>
        /// FTIR报告PDF
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public IActionResult ReplaceFTIRPdf([FromBody] FTIRInputDto parm)
        {
            _iqcService.ReplaceFTIRPdf(parm);

            return toResponse(StatusCodeType.Success, "FTIR报告生成成功!");
        }

        string GetColumnName(int columnNumber)
        {
            int dividend = columnNumber + 1;
            string columnName = "";
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }


        /// <summary>
        /// COC数据报告数据源
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public IActionResult COCDataSource([FromBody] COCInputDto parm) {

            if (string.IsNullOrEmpty(parm.COCID)) {
                return toResponse(StatusCodeType.Error, $"COCID不能为空！");
            }
            if (string.IsNullOrEmpty(parm.ID)) {
                return toResponse(StatusCodeType.Error, $"ID不能为空！");
            }
            if (parm.FIX_VALUE == null || parm.FIX_VALUE.Length == 0) {
                return toResponse(StatusCodeType.Error, $"FIX_VALUE不能为空！");
            }
            try {
                var ds = _iqcService.GetCOCDataSource(parm);
                string source = JsonConvert.SerializeObject(ds);
                return toResponse(StatusCodeType.Success, source);
            }
            catch (Exception ex) {
                return toResponse(StatusCodeType.Error, "COC数据获取失败:" + ex.ToString());
            }
        }

        /// <summary>
        /// COC数据报告API2
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public IActionResult COCReport([FromBody] COCInputDto parm) {

            if (string.IsNullOrEmpty(parm.COCID)) {
                return toResponse(StatusCodeType.Error, $"COCID不能为空！");
            }
            if (string.IsNullOrEmpty(parm.ID)) {
                return toResponse(StatusCodeType.Error, $"ID不能为空！");
            }
            if (parm.FIX_VALUE == null || parm.FIX_VALUE.Length == 0 ) {
                return toResponse(StatusCodeType.Error, $"FIX_VALUE不能为空！");
            }
            try {
                _iqcService.GetCOCfile(parm);
            }
            catch (Exception ex) {
                return toResponse(StatusCodeType.Error, "COC数据报告生成失败:" + ex.ToString());
            }
            return toResponse(StatusCodeType.Success, "COC数据报告生成成功！");
        }
    }
}
