using Meiam.System.Common.Helpers;
using Meiam.System.Model;
using Microsoft.AspNetCore.Mvc;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Meiam.System.Hostd.Controllers
{
    public class BaseController : ControllerBase
    {
        #region 统一返回封装

        /// <summary>
        /// 返回封装
        /// </summary>
        /// <param name="statusCode"></param>
        /// <returns></returns>
        public static JsonResult toResponse(StatusCodeType statusCode)
        {
            ApiResult response = new ApiResult();
            response.StatusCode = (int)statusCode;
            response.Message = statusCode.GetEnumText();

            return new JsonResult(response);
        }

        /// <summary>
        /// 返回封装
        /// </summary>
        /// <param name="statusCode"></param>
        /// <param name="retMessage"></param>
        /// <returns></returns>
        public static JsonResult toResponse(StatusCodeType statusCode, string retMessage)
        {
            ApiResult response = new ApiResult();
            response.StatusCode = (int)statusCode;
            response.Message = retMessage;

            return new JsonResult(response);
        }

        /// <summary>
        /// 返回封装
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="data"></param>
        /// <returns></returns>
        public static JsonResult toResponse<T>(T data)
        {
            ApiResult<T> response = new ApiResult<T>();
            response.StatusCode = (int)StatusCodeType.Success;
            response.Message = StatusCodeType.Success.GetEnumText();
            response.Data = data;
            return new JsonResult(response);
        }

        #endregion

        #region 常用方法函数
        public static string GetGUID
        {
            get
            {
                return SequentialGuid.Generate();
            }
        }

        #endregion

        #region  导出Excel相关
        protected virtual byte[] GetExcelContent(DataTable dataTable)
        {
            byte[] content = null;
            // 创建内存流
            using (var memoryStream = new MemoryStream())
            {
                // 创建工作簿
                IWorkbook workbook = new XSSFWorkbook();
                // 创建工作表
                ISheet sheet = workbook.CreateSheet("Sheet1");

                // 创建日期格式样式
                ICellStyle dateCellStyle = workbook.CreateCellStyle();
                IDataFormat dataFormat = workbook.CreateDataFormat();
                dateCellStyle.DataFormat = dataFormat.GetFormat("yyyy-MM-dd");

                // 创建表头样式（红色字体样式）
                ICellStyle redFontStyle = workbook.CreateCellStyle();
                IFont redFont = workbook.CreateFont();
                redFont.Color = IndexedColors.Red.Index; // 设置字体为红色
                redFontStyle.SetFont(redFont);
                redFontStyle.WrapText = true; // 启用换行
                // 默认表头样式（支持换行）
                ICellStyle defaultHeaderStyle = workbook.CreateCellStyle();
                defaultHeaderStyle.WrapText = true; // 启用换行
                // 创建表头
                IRow headerRow = sheet.CreateRow(0);
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    string colName = dataTable.Columns[i].ColumnName;
                    ICell cell = headerRow.CreateCell(i);
                    cell.SetCellValue(colName.TrimEnd('*'));
                    if (colName.EndsWith("*"))
                    {//必填字体为红色
                        cell.CellStyle = redFontStyle;
                    }
                    else
                    {
                        cell.CellStyle = defaultHeaderStyle;
                    }
                }
                // 填充数据
                for (int rowIndex = 0; rowIndex < dataTable.Rows.Count; rowIndex++)
                {
                    DataRow dataRow = dataTable.Rows[rowIndex];
                    IRow excelRow = sheet.CreateRow(rowIndex + 1);
                    for (int colIndex = 0; colIndex < dataTable.Columns.Count; colIndex++)
                    {
                        ICell cell = excelRow.CreateCell(colIndex);
                        // 获取当前列的数据类型
                        Type columnType = dataTable.Columns[colIndex].DataType;
                        if (columnType == typeof(DateTime))
                        {
                            // 如果是日期类型，设置格式为 yyyy-MM-dd
                            if (dataRow[colIndex] != DBNull.Value)
                            {
                                cell.SetCellValue((DateTime)dataRow[colIndex]);
                                cell.CellStyle = dateCellStyle;
                            }
                        }
                        else if (columnType == typeof(int) || columnType == typeof(double) || columnType == typeof(float))
                        {
                            // 如果是数值类型
                            if (dataRow[colIndex] != DBNull.Value)
                                cell.SetCellValue(Convert.ToDouble(dataRow[colIndex]));
                        }
                        else if (columnType == typeof(bool))
                        {
                            if (dataRow[colIndex].ToString() == "True")
                                cell.SetCellValue("是");
                            else
                                cell.SetCellValue("否");
                        }
                        else
                        {
                            // 其他类型（字符串等）
                            cell.SetCellValue(dataRow[colIndex]?.ToString() ?? string.Empty);
                        }
                    }
                }
                // 自动调整列宽
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    sheet.AutoSizeColumn(i);
                }
                // 写入内存流
                workbook.Write(memoryStream, true);
                // 不关闭流，确保我们可以后续使用它
                memoryStream.Flush(); // 确保流被刷新
                // 重置流位置
                memoryStream.Seek(0, SeekOrigin.Begin); // 推荐使用 Seek 替代 Position
                content = memoryStream.ToArray();
                memoryStream.Close();
            }
            return content;
        }

        #endregion

        /// <summary>
        /// 项目锁,防止并发问题
        /// </summary>
        public static readonly ConcurrentDictionary<string, object> _locks = new ConcurrentDictionary<string, object>();


        // 比较两个对象相同属性的方法
        protected bool CompareObjects(object obj1, object obj2)
        {
            // 获取类型
            var type1 = obj1.GetType();
            var type2 = obj2.GetType();

            // 获取公共属性
            var properties1 = type1.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            var properties2 = type2.GetProperties(BindingFlags.Public | BindingFlags.Instance);

            // 遍历属性
            foreach (var prop1 in properties1)
            {
                var prop2 = properties2.FirstOrDefault(p => p.Name == prop1.Name && p.PropertyType == prop1.PropertyType);
                if (prop2 != null)
                {
                    // 获取属性值
                    var value1 = prop1.GetValue(obj1);
                    var value2 = prop2.GetValue(obj2);

                    // 比较值
                    if (!Equals(value1, value2))
                    {
                        return true;
                    }
                }
            }

            return false;
        }
    }
}