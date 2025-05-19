﻿using DocumentFormat.OpenXml.Office2010.ExcelAc;
using Meiam.System.Common;
using NPOI.HPSF;
using NPOI.Util;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

public class ExcelHelper : IDisposable {
    private readonly ExcelPackage _excelPackage;
    private readonly FileInfo _fileInfo;
    private const int BaseRowHeight = 15; // 默认行高（磅）
    private const int BaseColumnWidth = 10; // 默认列宽（字符数）

    public ExcelHelper(string file) {
        ExcelPackage.License.SetNonCommercialPersonal("xiaodlll");
        _fileInfo = new FileInfo(file);
        _excelPackage = _fileInfo.Exists ?
            new ExcelPackage(_fileInfo) :
            new ExcelPackage();
    }

    public void AddTextToCell(string sheetName, string cellAddress, string text) {
        var worksheet = GetOrCreateWorksheet(sheetName);
        var cell = worksheet.Cells[cellAddress];

        // 定义 Unicode 上标字符映射表
        var superscriptMap = new Dictionary<char, string> {
        {'0', "⁰"}, {'1', "¹"}, {'2', "²"}, {'3', "³"}, {'4', "⁴"},
        {'5', "⁵"}, {'6', "⁶"}, {'7', "⁷"}, {'8', "⁸"}, {'9', "⁹"},
        {'+', "⁺"}, {'-', "⁻"}, {'=', "⁼"}, {'(', "⁽"}, {')', "⁾"}
    };

        // 处理 HTML 格式的数字的几次方，例如 5<sup>3</sup>
        int supStart = text.IndexOf("<sup>");
        while (supStart > 0) // 循环处理所有 <sup> 标签
        {
            // 查找前面的数字部分
            int numStart = supStart - 1;
            while (numStart >= 0 && char.IsDigit(text[numStart])) {
                numStart--;
            }
            numStart++; // 调整到数字开始位置

            int supEnd = text.IndexOf("</sup>", supStart);
            if (supEnd > supStart) {
                string baseNumber = text.Substring(numStart, supStart - numStart);
                string exponent = text.Substring(supStart + 5, supEnd - supStart - 5);

                // 转换指数为 Unicode 上标（不使用 LINQ）
                string superscript = "";
                foreach (char c in exponent) {
                    if (superscriptMap.ContainsKey(c)) {
                        superscript += superscriptMap[c];
                    }
                    else {
                        superscript += c;
                    }
                }

                // 替换 HTML 标签为 Unicode 上标
                text = text.Substring(0, numStart) + baseNumber + superscript +
                      (supEnd + 6 < text.Length ? text.Substring(supEnd + 6) : "");
            }

            // 继续查找下一个 <sup> 标签
            supStart = text.IndexOf("<sup>");
        }

        // 检查单元格是否已有数字格式
        if (IsNumericFormat(cell.Style.Numberformat.Format)) {
            // 尝试将文本转换为数字
            if (double.TryParse(text, out double numberValue)) {
                cell.Value = numberValue; // 设置为数字类型
                return;
            }
        }

        // 默认作为文本处理
        cell.Value = text;
    }

    private bool IsNumericFormat(string format) {
        // 判断是否为数字格式（简化版）
        if (string.IsNullOrEmpty(format)) return false;
        return format.Contains("0") || format.Contains("#") || format.Contains("?");
    }

    public void AddAttachsToCell(string sheetName, string cellAddress, string[] filePaths) {
        // 定义完整路径数组并初始化
        string[] fullFilePaths = new string[filePaths.Length];

        // 填充完整路径
        for (int i = 0; i < filePaths.Length; i++) {
            fullFilePaths[i] = Path.Combine(AppSettings.Configuration["AppSettings:FileServerPath"], filePaths[i].TrimStart('\\'));
        }

        var worksheet = GetOrCreateWorksheet(sheetName);
        var cell = worksheet.Cells[cellAddress];
        var startRow = cell.Start.Row;
        var startCol = cell.Start.Column;
        double cellHeight = GetMergedCellHeight(worksheet, startRow, startCol);
        double startLeft = 5;

        // 设置基础样式
        cell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
        worksheet.Row(startRow).CustomHeight = true;

        foreach (var filePath in fullFilePaths.Where(File.Exists)) {
            try {
                var (embedObject, isImage, width) = CreateEmbedObject(worksheet, filePath, cellHeight);

                var (actColumn, actStartLeft) = GetMergedCellLeft(worksheet, cell, startLeft);
                // 设置对象位置
                embedObject.SetPosition(
                    startRow - 1,
                    5,
                    actColumn - 1,
                    (int)(actStartLeft)
                );
                startLeft += width + 5;
            }
            catch (Exception ex) {
                throw new Exception($"文件[{filePath}]嵌入失败: {ex.ToString()}");
            }
        }
    }

    private (int, double) GetMergedCellLeft(ExcelWorksheet worksheet, ExcelRange cell, double startLeft) {
        // 检查单元格是否属于合并区域
        ExcelRangeBase mergedRange = null;
        foreach (var range in worksheet.MergedCells) {
            var merged = worksheet.Cells[range];
            if (merged.Start.Row <= cell.Start.Row && merged.End.Row >= cell.Start.Row &&
                merged.Start.Column <= cell.Start.Column && merged.End.Column >= cell.Start.Column) {
                mergedRange = merged;
                break;
            }
        }

        // 获取合并区域的起始列（如果有）
        int startColumn = mergedRange?.Start.Column ?? cell.Start.Column;

        // 计算合并区域内所有列的总宽度（像素）
        double mergedRegionWidth = 0;
        if (mergedRange != null) {
            for (int col = mergedRange.Start.Column; col <= mergedRange.End.Column; col++) {
                mergedRegionWidth += worksheet.Column(col).Width * 7; // 列宽转像素（近似值）
            }
        }
        else {
            // 非合并单元格，使用当前列宽
            mergedRegionWidth = worksheet.Column(startColumn).Width * 7;
        }

        // 计算startLeft之后的目标列和剩余偏移量
        double currentOffset = 0;
        int targetColumn = startColumn;

        // 累加列宽，直到超过startLeft
        while (currentOffset < startLeft && targetColumn <= worksheet.Dimension.End.Column) {
            double columnWidth = worksheet.Column(targetColumn).Width * 7;
            if (currentOffset + columnWidth > startLeft) {
                break; // 找到目标列
            }

            currentOffset += columnWidth;
            targetColumn++;
        }

        // 计算相对于目标列的剩余偏移量
        double remainingOffset = startLeft - currentOffset;

        return (targetColumn, remainingOffset);
    }

    /// <summary>
    /// 获取指定单元格的高度（如果是合并单元格，则返回合并区域的总高度）
    /// </summary>
    /// <param name="worksheet">工作表</param>
    /// <param name="row">行索引</param>
    /// <param name="column">列索引</param>
    /// <returns>单元格或合并区域的总高度（以磅为单位）</returns>
    private double GetMergedCellHeight(ExcelWorksheet worksheet, int row, int column) {
        // 检查单元格是否属于合并区域
        ExcelRangeBase mergedRange = null;
        foreach (var range in worksheet.MergedCells) {
            var merged = worksheet.Cells[range];
            if (merged.Start.Row <= row && merged.End.Row >= row &&
                merged.Start.Column <= column && merged.End.Column >= column) {
                mergedRange = merged;
                break;
            }
        }

        // 如果是合并单元格，计算合并区域的总高度
        if (mergedRange != null) {
            double totalHeight = 0;
            for (int r = mergedRange.Start.Row; r <= mergedRange.End.Row; r++) {
                totalHeight += worksheet.Row(r).Height; // 使用默认行高（如果未设置）
            }
            return totalHeight;
        }

        // 如果不是合并单元格，返回当前行的高度
        return worksheet.Row(row).Height;
    }

    public void InsertRows(string sheetName, int rowCount, int targetRow) {
        // 获取指定工作表
        var worksheet = GetOrCreateWorksheet(sheetName);
        // 在目标位置插入新行
        worksheet.InsertRow(targetRow, rowCount);
    }

    public void CopyRows(string sheetName, int[] sourceRows, int targetRow) {
        var worksheet = GetOrCreateWorksheet(sheetName);

        // 对源行进行降序排序，避免插入操作影响后续源行索引
        var sortedSourceRows = sourceRows.OrderByDescending(r => r).ToList();

        // 计算源行总高度（行数）
        int totalRowsToCopy = sortedSourceRows.Distinct().Count();

        // 插入足够的空行
        worksheet.InsertRow(targetRow, totalRowsToCopy);

        // 复制整个源区域到目标位置
        var sourceAddress = $"{sortedSourceRows.Min()}:{sortedSourceRows.Max()}";
        var targetAddress = $"{targetRow}:{targetRow + totalRowsToCopy - 1}";
        // 复制单元格内容和格式
        worksheet.Cells[sourceAddress].Copy(worksheet.Cells[targetAddress]);

        // 单独复制行高（解决行高问题的关键）
        for (int i = 0; i < sortedSourceRows.Count; i++) {
            int sourceRowIndex = sortedSourceRows[i];
            int targetRowIndex = targetRow + i;

            // 复制行高
            worksheet.Row(targetRowIndex).Height = worksheet.Row(sourceRowIndex).Height;
        }
    }

    public void CopyCells(string sheetName, string cellAddress, string targetCells) {
        // 获取指定工作表
        var worksheet = GetOrCreateWorksheet(sheetName);

        // 解析源单元格范围
        var sourceRange = worksheet.Cells[cellAddress];
        if (sourceRange == null) {
            throw new ArgumentException($"源单元格地址 '{cellAddress}' 无效。");
        }

        // 解析目标单元格位置
        var targetRange = worksheet.Cells[targetCells];
        if (targetRange == null) {
            throw new ArgumentException($"目标单元格地址 '{targetCells}' 无效。");
        }

        // 复制单元格内容、格式和公式
        sourceRange.Copy(targetRange);

        // 如果源范围和目标范围尺寸不同，调整目标区域大小以匹配源区域
        if (sourceRange.Count() != targetRange.Count()) {
            var sourceDimension = sourceRange.GetEnumerator().Current.Address;
            var targetDimension = targetRange.GetEnumerator().Current.Address;

            // 计算源区域的行列数
            int sourceRows = GetExcelRows(sourceDimension);
            int sourceCols = GetExcelColumns(sourceDimension);

            // 扩展目标区域以匹配源区域大小
            var expandedTargetRange = worksheet.Cells[
                targetRange.Start.Row,
                targetRange.Start.Column,
                targetRange.Start.Row + sourceRows - 1,
                targetRange.Start.Column + sourceCols - 1
            ];

            // 重新应用复制操作到扩展后的目标区域
            sourceRange.Copy(expandedTargetRange);
        }
    }

    public void CopySheet(string sourceExcelPath, string sourceSheetName, string targetSheetName) {
        // 1. 验证源文件是否存在
        if (!File.Exists(sourceExcelPath))
            throw new FileNotFoundException("源 Excel 文件不存在", sourceExcelPath);

        // 2. 加载源 Excel 文件和工作表
        using (var sourcePackage = new ExcelPackage(new FileInfo(sourceExcelPath))) {
            var sourceSheet = sourcePackage.Workbook.Worksheets[sourceSheetName];
            if (sourceSheet == null)
                throw new ArgumentException($"源工作表 '{sourceSheetName}' 不存在");

            // 3. 确定目标工作表的原有位置（若存在）
            var targetSheet = _excelPackage.Workbook.Worksheets[targetSheetName];
            int originalPosition = targetSheet?.Index ?? _excelPackage.Workbook.Worksheets.Count;

            // 4. 删除已存在的目标工作表（避免重名冲突）
            if (targetSheet != null)
                _excelPackage.Workbook.Worksheets.Delete(targetSheet);

            // 5. 创建新的目标工作表（临时位置在末尾）
            var newTargetSheet = _excelPackage.Workbook.Worksheets.Add(targetSheetName);

            // 6. 复制源工作表的全部内容（含样式、公式等）
            if (sourceSheet.Dimension != null)  // 源工作表非空时复制
            {
                var sourceRange = sourceSheet.Cells[sourceSheet.Dimension.Address];
                var targetRange = newTargetSheet.Cells[sourceSheet.Dimension.Address];
                sourceRange.Copy(targetRange);  // 关键：内置范围复制（保留所有样式）
            }

            // 7. 调整目标工作表到原有位置
            if(originalPosition == 0) {
                _excelPackage.Workbook.Worksheets.MoveToStart(targetSheetName);
            }
            else {
                _excelPackage.Workbook.Worksheets.MoveBefore(targetSheetName, _excelPackage.Workbook.Worksheets[originalPosition].Name);
            }
        }
    }

    /// <summary>
    /// 获取Excel区域的行数
    /// </summary>
    private int GetExcelRows(string rangeAddress) {
        var address = new ExcelAddress(rangeAddress);
        return address.End.Row - address.Start.Row + 1;
    }

    /// <summary>
    /// 获取Excel区域的列数
    /// </summary>
    private int GetExcelColumns(string rangeAddress) {
        var address = new ExcelAddress(rangeAddress);
        return address.End.Column - address.Start.Column + 1;
    }

    public void Dispose() {
        try {
            if (!_excelPackage.Workbook.Worksheets.Any()) {
                _excelPackage.Workbook.Worksheets.Add("Sheet1");
            }
            _excelPackage.SaveAs(_fileInfo);
        }
        finally {
            _excelPackage?.Dispose();
        }
    }

    private ExcelWorksheet GetOrCreateWorksheet(string sheetName) {
        return _excelPackage.Workbook.Worksheets[sheetName] ??
            _excelPackage.Workbook.Worksheets.Add(sheetName);
    }

    private (ExcelDrawing embedObject, bool isImage, int width) CreateEmbedObject(
        ExcelWorksheet worksheet,
        string filePath ,double targetRowHeight) {
        var extension = Path.GetExtension(filePath).ToLower();
        var imageTypes = new[] { ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".pdf" };
        int width = (int)(targetRowHeight);
        if (imageTypes.Contains(extension)) {
            //pdf
            if (extension.Contains(".pdf")) {
                using (var stream = FileToImageHelper.ConvertFirstPageToImageStream(filePath)) {
                    var picture = worksheet.Drawings.AddPicture(
                      Guid.NewGuid().ToString(),
                      stream
                    );
                    if (imageTypes.Contains(extension)) {
                        // 添加图片后立即缩放
                        width = ScaleImageToCell(picture, targetRowHeight);
                    }
                    return (picture, true, width);
                }
            }
            else {
                byte[] imageBytes = File.ReadAllBytes(filePath);
                using (var stream = new MemoryStream(imageBytes)) {
                    var picture = worksheet.Drawings.AddPicture(
                      Guid.NewGuid().ToString(),
                      stream
                    );
                    if (imageTypes.Contains(extension)) {
                        // 添加图片后立即缩放
                        width = ScaleImageToCell(picture, targetRowHeight);
                    }
                    return (picture, true, width);
                }
            }
        }
        else {
            var oleObject = worksheet.Drawings.AddOleObject(
                Guid.NewGuid().ToString(),
                filePath
            );
            oleObject.SetSize(width, width);
            return (oleObject, false, width);
        }
    }
    private int ScaleImageToCell(ExcelPicture picture, double cellHeightPoints) {
        // 精确单位转换（点→像素）
        double cellHeightPixels = (cellHeightPoints - 5) * 96 / 72; // 96 DPI标准

        // 等比缩放计算
        double scale = cellHeightPixels / picture.Image.Bounds.Height;
        int newWidth = (int)(picture.Image.Bounds.Width * scale);

        // 设置最终尺寸
        picture.SetSize(newWidth, (int)cellHeightPixels);
        return newWidth;
    }

}