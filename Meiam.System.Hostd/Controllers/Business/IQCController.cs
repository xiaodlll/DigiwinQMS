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
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
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

            _iqcService.GetInspectReport(parm);

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

            _iqcService.GetCPKfile(parm.INSPECT_DEV2ID, parm.UserName);
            return toResponse(StatusCodeType.Success, "CPK数据报告生成成功！");

            //byte[] fileContents = _iqcService.getcpk(listToSave);


            //byte[] fileContents = _iqcService.GetCPKReport(parm);
            //byte[] fileContents = null;
            //// 3. 返回文件流
            //var fileName = $"零件清单_{DateTime.Now:yyyyMMdd}.jpg";
            //return File(fileContents, "image/jpeg", fileName);
        }

        /// <summary>
        /// CPK数据报告测试Excel
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public IActionResult TestCPKExcel()
        {
            // 创建工作簿
            IWorkbook workbook = new XSSFWorkbook();
            // 创建工作表
            ISheet sheet = workbook.CreateSheet("Sheet1");

            // 定义表头字体样式
            ICellStyle headerCellStyle = workbook.CreateCellStyle();
            IFont headerFont = workbook.CreateFont();
            headerFont.FontHeightInPoints = 12;
            headerFont.IsBold = true;
            headerCellStyle.SetFont(headerFont);
            // 设置表头边框样式
            headerCellStyle.BorderTop = BorderStyle.Thin;
            headerCellStyle.BorderBottom = BorderStyle.Thin;
            headerCellStyle.BorderLeft = BorderStyle.Thin;
            headerCellStyle.BorderRight = BorderStyle.Thin;
            headerCellStyle.TopBorderColor = IndexedColors.Black.Index;
            headerCellStyle.BottomBorderColor = IndexedColors.Black.Index;
            headerCellStyle.LeftBorderColor = IndexedColors.Black.Index;
            headerCellStyle.RightBorderColor = IndexedColors.Black.Index;

            // 定义普通单元格边框样式
            ICellStyle normalCellStyle = workbook.CreateCellStyle();
            normalCellStyle.BorderTop = BorderStyle.Thin;
            normalCellStyle.BorderBottom = BorderStyle.Thin;
            normalCellStyle.BorderLeft = BorderStyle.Thin;
            normalCellStyle.BorderRight = BorderStyle.Thin;
            normalCellStyle.TopBorderColor = IndexedColors.Black.Index;
            normalCellStyle.BottomBorderColor = IndexedColors.Black.Index;
            normalCellStyle.LeftBorderColor = IndexedColors.Black.Index;
            normalCellStyle.RightBorderColor = IndexedColors.Black.Index;

            // 定义数字格式单元格样式
            ICellStyle numberCellStyle = workbook.CreateCellStyle();
            IDataFormat dataFormat = workbook.CreateDataFormat();
            numberCellStyle.DataFormat = dataFormat.GetFormat("#,##0.0##"); // 例如 1,234.56
            numberCellStyle.BorderTop = BorderStyle.Thin;
            numberCellStyle.BorderBottom = BorderStyle.Thin;
            numberCellStyle.BorderLeft = BorderStyle.Thin;
            numberCellStyle.BorderRight = BorderStyle.Thin;
            numberCellStyle.TopBorderColor = IndexedColors.Black.Index;
            numberCellStyle.BottomBorderColor = IndexedColors.Black.Index;
            numberCellStyle.LeftBorderColor = IndexedColors.Black.Index;
            numberCellStyle.RightBorderColor = IndexedColors.Black.Index;

            // 定义百分比格式单元格样式
            ICellStyle percentCellStyle = workbook.CreateCellStyle();
            IDataFormat dataFormat1 = workbook.CreateDataFormat();
            percentCellStyle.DataFormat = dataFormat1.GetFormat("0.0#%"); // 
            percentCellStyle.BorderTop = BorderStyle.Thin;
            percentCellStyle.BorderBottom = BorderStyle.Thin;
            percentCellStyle.BorderLeft = BorderStyle.Thin;
            percentCellStyle.BorderRight = BorderStyle.Thin;
            percentCellStyle.TopBorderColor = IndexedColors.Black.Index;
            percentCellStyle.BottomBorderColor = IndexedColors.Black.Index;
            percentCellStyle.LeftBorderColor = IndexedColors.Black.Index;
            percentCellStyle.RightBorderColor = IndexedColors.Black.Index;
            

            // 写入表头信息
            string[] headers = {
            "FAI&CPK Data Sheet - Rev 06", "Unnamed: 1", "Unnamed: 2", "RoHS  HF",
            "Unnamed: 4", "Unnamed: 5", "Unnamed: 6", "Unnamed: 7",
            "Unnamed: 8", "Unnamed: 9", "Unnamed: 10", "Unnamed: 11",
            "Unnamed: 12", "Unnamed: 13", "Yiled<90%", "Unnamed: 15",
            "Unnamed: 16", "Unnamed: 17"
        };

            int rowIndex = 0;
            IRow headerRow = sheet.CreateRow(rowIndex++);
            for (int i = 0; i < headers.Length; i++)
            {
                ICell cell = headerRow.CreateCell(i);
                cell.SetCellValue(headers[i]);
                cell.CellStyle = headerCellStyle;
            }

            // 写入前几行的静态数据
            string[][] staticData = {
            new string[] { "Part Number :", "RGPZ - 322J ADH - P", null, null, "Revision :", "V1", "Supplier :", "Jiutai", null, "Inspector:", "张钰俊" },
            new string[] { "Part Description :", "PSA", null, null, null, null, "Cavity / Tool # :", null, "50709", "Date:", "2025 - 01 - 16 00:00:00" },
            new string[] { "Request Process" },
            new string[] { "Dimension Description" },
            new string[] { "Comments" },
            new string[] { "Proposed +Tol" },
            new string[] { "Proposed -Tol" },
            new string[] { "Mean Shift Amount" },
            new string[] { "Changed Mean" },
            new string[] { "Proposed USL" },
            new string[] { "Proposed LSL" },
            new string[] { "Proposed Yield" },
            new string[] { "Distribution Type", "DoubleSides", "DoubleSides", "DoubleSides", "DoubleSides", "DoubleSides", "DoubleSides", "DoubleSides", "DoubleSides", "DoubleSides", "DoubleSides"},
        };
            for (int i = 0; i < staticData.Length; i++)
            {
                IRow row = sheet.CreateRow(rowIndex++);
                for (int j = 0; j < staticData[i].Length; j++)
                {
                    ICell cell = row.CreateCell(j);
                    if (staticData[i][j] != null)
                    {
                        cell.SetCellValue(staticData[i][j]);
                        cell.CellStyle = normalCellStyle;
                    }
                }
            }

            // 写入 14 - 19 行的数据（DataTable1）
            string[][] dataTable1 = {
            new string[] { "Dim. No.", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "57115HSB（完成品）", " 57115HSB （完成品）", "IDBS025D(75)B（完成品）", "57115HSB（原材）", "57115HSB （原材）" },
            new string[] { "Nominal Dim.", "0.15", "8.44", "8.74", "0.15", "57.38", "70.54", "57.23", "2", "2", "2", "2", "81.52", "0.15", "0.15", "0.077", "0.15", "0.15" },
            new string[] { "Dim Model", "DC", "DC", "SPC", "DC", "DC", "SPC", "DC", "-", "-", "-", "-", "-", "万分尺", "厚度仪减去千分尺", "厚度仪", "万分尺", "厚度仪减去千分尺" },
            new string[] { "Tol. Max. (+)", "0.12", "0.12", "0.15", "0.12", "0.12", "0.15", "0.12", "0.15", "0.15", "0.15", "0.15", "1", "0.01", "0.01", "0.003", "0.01", "0.01" },
            new string[] { "Tol. Min. (-)", "0.12", "0.12", "0.15", "0.12", "0.15", "0.15", "0.12", "0.15", "0.15", "0.15", "0.15", "1", "0.03", "0.03", "0.003", "0.03", "0.03" },
        };
            for (int i = 0; i < dataTable1.Length; i++)
            {
                IRow row = sheet.CreateRow(rowIndex++);
                for (int j = 0; j < dataTable1[i].Length; j++)
                {
                    ICell cell = row.CreateCell(j);
                    cell.SetCellValue(dataTable1[i][j]);
                    cell.CellStyle = normalCellStyle;
                }
            }

            // 写入 20 - 31 行的公式
            string[][] formulasData = {
            new string[] { "USL", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
            new string[] { "LSL", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
            new string[] { "Std Dev", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
            new string[] { "Mean", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
            new string[] { "Maximum", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
            new string[] { "Minimum", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
            new string[] { "Cp", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
            new string[] { "Cpkl", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
            new string[] { "Cpku", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
            new string[] { "Cpk", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
            new string[] { "Projected Yields", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
            new string[] { "Mean Drift", null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null },
        };
            for (int i = 0; i < formulasData.Length; i++)
            {
                IRow row = sheet.CreateRow(rowIndex++);
                for (int j = 0; j < formulasData[i].Length; j++)
                {
                    ICell cell = row.CreateCell(j);
                    cell.SetCellValue(formulasData[i][j]);
                    cell.CellStyle = normalCellStyle;
                }
            }
            // 假设的公式示例，你需要根据实际情况修改
            for (int col = 1; col < formulasData[0].Length; col++)
            {
                string colName = GetColumnName(col);
                // USL
                sheet.GetRow(19).GetCell(col).CellFormula = $"{colName}16+ABS({colName}18)";
                // LSL
                sheet.GetRow(20).GetCell(col).CellFormula = $"{colName}16-ABS({colName}19)";
                // Std Dev
                sheet.GetRow(21).GetCell(col).CellFormula = $"STDEV({colName}32:{colName}65194)";
                // Mean
                sheet.GetRow(22).GetCell(col).CellFormula = $"AVERAGE({colName}32:{colName}65194)";
                // Maximum
                sheet.GetRow(23).GetCell(col).CellFormula = $"MAX({colName}32:{colName}65194)";
                // Minimum
                sheet.GetRow(24).GetCell(col).CellFormula = $"MIN(({colName}20-{colName}23)/(3*{colName}22),({colName}23-{colName}21)/(3*{colName}22))";
                // Cp
                sheet.GetRow(25).GetCell(col).CellFormula = $"(({colName}20)-({colName}21))/(6*{colName}22)";
                // Cpkl
                sheet.GetRow(26).GetCell(col).CellFormula = $"({colName}23-{colName}21)/(3*{colName}22)";
                // Cpku
                sheet.GetRow(27).GetCell(col).CellFormula = $"({colName}20-{colName}23)/(3*{colName}22)";
                // Cpk
                sheet.GetRow(28).GetCell(col).CellFormula = $"MIN(({colName}20-{colName}23)/(3*{colName}22),({colName}23-{colName}21)/(3*{colName}22))";
                // Projected Yields
                sheet.GetRow(29).GetCell(col).CellFormula = $"IF({colName}14=\"DoubleSides\",NORMSDIST(({colName}20-{colName}23)/{colName}22)-NORMSDIST(({colName}21-{colName}23)/{colName}22),IF({colName}14=\"SingleSide-USL\",NORMSDIST(({colName}20-{colName}23)/{colName}22),IF({colName}14=\"SingleSide-LSL\",1-NORMSDIST((B21-B23)/B22),IF({colName}14=\"Actual Yield (n>=100)\",IF(COUNT({colName}32:{colName}65194)>=100,(COUNTIF({colName}32:{colName}65194,\"<=\"&B20)-COUNTIF({colName}32:{colName}65194,\"<\"&{colName}21))/COUNT({colName}32:{colName}65194),\"Too Few Records\"),\"Error\"))))";
                // Mean Drift
                sheet.GetRow(30).GetCell(col).CellFormula = $"{colName}23-({colName}20+{colName}21)/2";
            }

            for (int i = 0; i < formulasData.Length; i++)
            {
                IRow row = sheet.GetRow(i + 19);
                for (int j = 0; j < formulasData[i].Length; j++)
                {
                    ICell cell = row.GetCell(j);
                    if (i + 19 == 29)
                    {
                        cell.CellStyle = percentCellStyle;
                    }
                    else
                    {
                        cell.CellStyle = numberCellStyle;
                    }
                }
            }

            // 写入 32 - 63 行的数据（DataTable2）
            string[][] dataTable2 = {
            new string[] { "1", "0.184", "8.44", "8.812", "0.143", "57.359", "70.514", "57.216", "2.026", "2.032", "1.983", "2.016", "81.471", "0.1445", "0.1474", "0.0766", "0.1476", "0.1478" },
            new string[] { "2", "0.17", "8.434", "8.802", "0.15", "57.373", "70.531", "57.223", "1.995", "2.022", "1.995", "2.069", "81.48", "0.1371", "0.1517", "0.0775", "0.148", "0.1486" },
            new string[] { "3", "0.184", "8.44", "8.812", "0.143", "57.359", "70.514", "57.216", "2.026", "2.032", "1.983", "2.016", "81.471", "0.1445", "0.1474", "0.0766", "0.1476", "0.1478" },
            new string[] { "4", "0.17", "8.434", "8.802", "0.15", "57.373", "70.531", "57.223", "1.995", "2.022", "1.995", "2.069", "81.48", "0.1371", "0.1517", "0.0775", "0.148", "0.1486" },
            new string[] { "5", "0.169", "8.438", "8.808", "0.152", "57.378", "70.538", "57.226", "2.001", "2.035", "1.976", "2.048", "81.468", "0.1444", "0.1466", "0.0777", "0.1488", "0.1517" },
            new string[] { "6", "0.158", "8.453", "8.78", "0.14", "57.373", "70.519", "57.233", "2.021", "2.061", "1.991", "2.061", "81.611", "0.1449", "0.1473", "0.0769", "0.1519", "0.1492" },
            new string[] { "7", "0.168", "8.441", "8.793", "0.144", "57.373", "70.526", "57.229", "2.006", "2.039", "2.007", "2.055", "81.602", "0.1451", "0.1484", "0.0772", "0.1522", "0.1466" },
            new string[] { "8", "0.158", "8.454", "8.785", "0.137", "57.361", "70.517", "57.223", "1.995", "2.029", "1.983", "2.049", "81.492", "0.1425", "0.1513", "0.077", "0.1463", "0.1484" },
            new string[] { "9", "0.165", "8.451", "8.781", "0.137", "57.352", "70.519", "57.215", "1.999", "2.025", "1.993", "2.048", "81.617", "0.1432", "0.1534", "0.0771", "0.1509", "0.1519" },
            new string[] { "10", "0.181", "8.449", "8.795", "0.148", "57.368", "70.535", "57.22", "2.01", "2.01", "1.974", "2.061", "81.521", "0.137", "0.1497", "0.0768", "0.147", "0.1499" },
            new string[] { "11", "0.183", "8.45", "8.791", "0.143", "57.353", "70.527", "57.21", "1.996", "2.038", "1.999", "2.046", "81.537", "0.1359", "0.1497", "0.0768", "0.1508", "0.1489" },
            new string[] { "12", "0.196", "8.445", "8.788", "0.135", "57.356", "70.514", "57.221", "2.021", "2.038", "2.001", "2.076", "81.547", "0.1415", "0.1512", "0.0767", "0.1518", "0.1522" },
            new string[] { "13", "0.206", "8.442", "8.792", "0.141", "57.354", "70.515", "57.213", "2.011", "2.034", "1.99", "2.042", "81.544", "0.137", "0.1534", "0.0769", "0.1529", "0.1491" },
            new string[] { "14", "0.212", "8.436", "8.787", "0.135", "57.35", "70.522", "57.215", "2.02", "2.005", "2.015", "2.027", "81.49", "0.1423", "0.1477", "0.0775", "0.1465", "0.1486" },
            new string[] { "15", "0.17", "8.448", "8.797", "0.128", "57.351", "70.522", "57.223", "2.016", "2.061", "1.991", "2.037", "81.583", "0.1417", "0.1532", "0.0776", "0.1485", "0.14" }
        };
            for (int i = 0; i < dataTable2.Length; i++)
            {
                IRow row = sheet.CreateRow(rowIndex++);
                for (int j = 0; j < dataTable2[i].Length; j++)
                {
                    ICell cell = row.CreateCell(j);
                    cell.SetCellValue(double.Parse(dataTable2[i][j]));
                    cell.CellStyle = numberCellStyle;
                }
            }

            byte[] fileContents;
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Write(ms);
                fileContents = ms.ToArray(); // 将流转换为 byte[]
            }
            // 3. 返回文件流
            var fileName = $"Excel报告.xlsx";
            return File(fileContents, "Excel/xlsx", fileName);
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
    }
}
