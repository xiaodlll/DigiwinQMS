using DocumentFormat.OpenXml.Spreadsheet;
using ICSharpCode.SharpZipLib.Zip;
using iText.Barcodes.Dmcode;
using iText.Layout.Font;
using Meiam.System.Common;
using Microsoft.IdentityModel.Tokens;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.OleObject;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Xml;

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
        int caretIndex = text.IndexOf('^');
        while (caretIndex != -1) {
            // 定义上标终止分隔符集合
            char[] separators = { ';', ',', '-', '\\', '/', '~' };
            int start = caretIndex + 1;

            // 若 "^" 位于字符串末尾，直接移除
            if (start >= text.Length) {
                text = text.Remove(caretIndex, 1);
                caretIndex = text.IndexOf('^');
                continue;
            }

            // 查找上标内容的终止位置（第一个分隔符或字符串结尾）
            int separatorIndex = text.IndexOfAny(separators, start);
            int end = separatorIndex != -1 ? separatorIndex : text.Length;

            // 提取需要转换为上标的文本
            string supText = text.Substring(start, end - start);

            // 转换为 Unicode 上标字符
            string superscript = "";
            foreach (char c in supText) {
                if (superscriptMap.ContainsKey(c)) {
                    superscript += superscriptMap[c];
                }
                else {
                    superscript += c; // 非映射字符保留原样
                }
            }

            // 重构字符串（移除 "^" 并替换为上标内容）
            text = text.Substring(0, caretIndex) + superscript + text.Substring(end);

            // 查找下一个 "^" 继续处理
            caretIndex = text.IndexOf('^');
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

    /// <summary>
    /// 解析包含^的文本，分离出普通文本和上标文本
    /// </summary>
    /// <param name="text">输入文本</param>
    /// <returns>包含(是否上标, 文本内容)的列表</returns>
    private List<(bool IsSuperscript, string Text)> ParseTextWithSuperscript(string text) {
        var parts = new List<(bool, string)>();
        if (string.IsNullOrEmpty(text))
            return parts;

        // 定义上标结束的分隔符集合
        char[] superscriptSeparators = { ';', ',', '-', '\\', '/', '~' };
        int currentIndex = 0;
        int textLength = text.Length;

        while (currentIndex < textLength) {
            // 查找下一个^的位置
            int caretIndex = text.IndexOf('^', currentIndex);
            if (caretIndex == -1) {
                // 没有^了，剩余内容作为普通文本
                parts.Add((false, text.Substring(currentIndex)));
                break;
            }

            // 处理^之前的普通文本
            if (caretIndex > currentIndex) {
                string normalText = text.Substring(currentIndex, caretIndex - currentIndex);
                parts.Add((false, normalText));
            }

            // 处理^后面的上标文本
            int superscriptStart = caretIndex + 1;
            if (superscriptStart >= textLength) {
                // ^是最后一个字符，上标文本为空
                parts.Add((true, string.Empty));
                currentIndex = textLength;
                break;
            }

            // 查找上标后的第一个分隔符
            int separatorIndex = text.IndexOfAny(superscriptSeparators, superscriptStart);
            if (separatorIndex == -1) {
                // 没有分隔符，^右侧所有内容为上标
                string superscript = text.Substring(superscriptStart);
                parts.Add((true, superscript));
                currentIndex = textLength;
                break;
            }
            else {
                // 有分隔符，^到分隔符之间的内容为上标
                if (separatorIndex > superscriptStart) {
                    string superscript = text.Substring(superscriptStart, separatorIndex - superscriptStart);
                    parts.Add((true, superscript));
                }
                else {
                    // 分隔符紧跟^，上标文本为空
                    parts.Add((true, string.Empty));
                }
                currentIndex = separatorIndex; // 从分隔符位置继续解析
            }
        }

        return parts;
    }

    private bool IsNumericFormat(string format) {
        // 判断是否为数字格式（简化版）
        if (string.IsNullOrEmpty(format)) return false;
        return format.Contains("0") || format.Contains("#") || format.Contains("?");
    }

    public void AddAttachsToCell(string sheetName, string cellAddress, ExcelAttechFile[] excelAttechFiles, bool attMode = false, bool attConvertPic = false) {

        if (excelAttechFiles.Length == 0)
            return;

        // 定义完整路径数组并初始化
        // 填充完整路径
        for (int i = 0; i < excelAttechFiles.Length; i++) {
            excelAttechFiles[i].FilePath = Path.Combine(AppSettings.Configuration["AppSettings:FileServerPath"], excelAttechFiles[i].FilePath.TrimStart('\\'));
            if (string.IsNullOrEmpty(excelAttechFiles[i].FileName)) {
                excelAttechFiles[i].FileName = Path.GetFileName(excelAttechFiles[i].FilePath);
            }
        }

        if ((excelAttechFiles[0].FilePath.EndsWith(".xlsx") || excelAttechFiles[0].FilePath.EndsWith(".xls")) ) {
            if (!attMode) {
                if (!attConvertPic) {
                    CopySheet(excelAttechFiles, sheetName, cellAddress);
                    return;
                }
            }
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
        int countOnject = 0;
        foreach (var excelAttechFile in excelAttechFiles) {
            try {
                var (embedObjects, isImage, width) = CreateEmbedObject(worksheet, excelAttechFile, cellHeight, attMode);
                foreach (var embedObject in embedObjects) {
                    var (actColumn, actStartLeft) = GetMergedCellLeft(worksheet, cell, startLeft);
                    embedObject.SetPosition(startRow - 1, 5, actColumn - 1, (int)(actStartLeft));

                    startLeft += width + 5;
                    countOnject++;
                }
            }
            catch (Exception ex) {
                throw new Exception($"文件[{excelAttechFile.FilePath}]嵌入失败: {ex.ToString()}");
            }
        }
    }

    private (int, double) GetMergedCellLeft(ExcelWorksheet worksheet, ExcelRange cell, double startLeft) {
        // 检查单元格是否属于合并区域
        ExcelRangeBase mergedRange = null;
        foreach (var range in worksheet.MergedCells) {
            if (string.IsNullOrEmpty(range))
                continue;
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
            if (string.IsNullOrEmpty(range))
                continue;
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

    #region CopyRows
    /// <summary>
    /// 安全版本 - 处理包含合并单元格的情况
    /// </summary>
    public void CopyRows(string sheetName, string cellsZone, int copyRows) {
        if (copyRows <= 0) return;

        var worksheet = GetOrCreateWorksheet(sheetName);
        var address = new ExcelAddress(cellsZone);

        int startRow = address.Start.Row;
        int startCol = address.Start.Column;
        int endRow = address.End.Row;
        //int endCol = address.End.Column;
        int maxColumn = worksheet.Dimension.End.Column;
        int endCol = maxColumn;

        int sourceRowCount = endRow - startRow + 1;

        try {
            //优先使用快速复制方法
            CopyRowsSimple(worksheet, startRow, endRow, startCol, endCol, copyRows);
        }
        catch {
            //使用安全的逐行复制方法
            CopyRowsWithMergedCells(worksheet, startRow, endRow, startCol, endCol, copyRows);
        }
    }

    /// <summary>
    /// 逐行复制（保留格式）
    /// </summary>
    private void CopyRowsSimple(ExcelWorksheet worksheet, int startRow, int endRow, int startCol, int endCol, int copyRows) {
        int sourceRowCount = endRow - startRow + 1; // 源区域行数
        int totalRowsToCopy = sourceRowCount * copyRows; // 总需复制的行数

        // 先在目标位置插入空白行（为复制内容预留空间）
        int targetStartRow = endRow + 1;
        worksheet.InsertRow(targetStartRow, totalRowsToCopy);

        // 逐行复制（避免批量复制导致的格式丢失）
        for (int copyIndex = 0; copyIndex < copyRows; copyIndex++) {
            // 循环复制源区域的每一行
            for (int rowIndex = 0; rowIndex < sourceRowCount; rowIndex++) {
                int sourceRow = startRow + rowIndex; // 当前源行
                                                     // 计算当前目标行（按复制次数和源行索引偏移）
                int targetRow = targetStartRow + copyIndex * sourceRowCount + rowIndex;

                // 复制当前行的单元格内容和格式（精确到列范围）
                var sourceCellRange = worksheet.Cells[sourceRow, startCol, sourceRow, endCol];
                var targetCellRange = worksheet.Cells[targetRow, startCol, targetRow, endCol];
                sourceCellRange.Copy(targetCellRange);

                // 复制行高（保持与源行一致）
                worksheet.Row(targetRow).Height = worksheet.Row(sourceRow).Height;
            }
        }
    }

    /// <summary>
    /// 安全复制 - 用于包含垂直合并单元格的情况 (适配新版 EPPlus)
    /// </summary>
    private void CopyRowsWithMergedCells(ExcelWorksheet worksheet, int startRow, int endRow, int startCol, int endCol, int copyRows) {
        int sourceRowCount = endRow - startRow + 1;

        // 预先收集源区域中的所有合并单元格 [citation:6]
        var mergedRanges = new List<ExcelRangeBase>();
        foreach (var mergedRange in worksheet.MergedCells) {
            if (string.IsNullOrEmpty(mergedRange))
                continue;
            ExcelAddressBase address = new ExcelAddressBase(mergedRange);
            // 检查合并区域是否与源区域相交
            if (address.Start.Row <= endRow && address.End.Row >= startRow &&
                address.Start.Column <= endCol && address.End.Column >= startCol) {
                mergedRanges.Add(worksheet.Cells[mergedRange]);
            }
        }

        // 逐批次复制
        for (int i = 0; i < copyRows; i++) {
            int targetStartRow = endRow + 1 + (i * sourceRowCount);

            // 插入目标行 [citation:3]
            worksheet.InsertRow(targetStartRow, sourceRowCount);

            // 复制单元格内容、样式和公式
            for (int rowOffset = 0; rowOffset < sourceRowCount; rowOffset++) {
                int sourceRow = startRow + rowOffset;
                int targetRow = targetStartRow + rowOffset;

                // 复制行高和隐藏状态
                worksheet.Row(targetRow).Height = worksheet.Row(sourceRow).Height;
                worksheet.Row(targetRow).Hidden = worksheet.Row(sourceRow).Hidden;

                // 逐列复制单元格
                for (int col = startCol; col <= endCol; col++) {
                    var sourceCell = worksheet.Cells[sourceRow, col];
                    var targetCell = worksheet.Cells[targetRow, col];

                    // 只复制非合并单元格或合并区域的第一个单元格 [citation:6]
                    if (!sourceCell.Merge || IsFirstCellInMerge(worksheet, sourceRow, col)) {
                        // 使用 Copy 方法复制单元格，但需要注意其在新版 EPPlus 中的行为 [citation:5]
                        sourceCell.Copy(targetCell);

                        // 显式复制一些关键样式属性以确保兼容性 [citation:2][citation:5]
                        CopyCellStyle(sourceCell, targetCell);
                    }
                }
            }

            // 在目标区域重新创建合并单元格 [citation:3]
            foreach (var mergedRange in mergedRanges) {
                // 计算源合并区域在目标工作表中的新位置
                int newStartRow = mergedRange.Start.Row - startRow + targetStartRow;
                int newEndRow = mergedRange.End.Row - startRow + targetStartRow;
                int newStartCol = mergedRange.Start.Column;
                int newEndCol = mergedRange.End.Column;

                // 确保新位置在目标行范围内
                if (newStartRow >= targetStartRow && newEndRow < targetStartRow + sourceRowCount) {
                    // 创建新的合并单元格 [citation:8]
                    worksheet.Cells[newStartRow, newStartCol, newEndRow, newEndCol].Merge = true;

                    // 确保合并区域的样式与源区域第一个单元格一致 [citation:4]
                    CopyCellStyle(mergedRange, worksheet.Cells[newStartRow, newStartCol, newEndRow, newEndCol]);
                }
            }
        }
    }

    /// <summary>
    /// 检查单元格是否是合并区域的第一个单元格 [citation:6]
    /// </summary>
    private bool IsFirstCellInMerge(ExcelWorksheet worksheet, int row, int col) {
        var cell = worksheet.Cells[row, col];
        if (!cell.Merge) return false;

        var mergeAddress = worksheet.MergedCells[row, col];
        if (mergeAddress == null) return false;

        ExcelAddress address = new ExcelAddress(mergeAddress);
        return address.Start.Row == row && address.Start.Column == col;
    }

    /// <summary>
    /// 复制单元格样式 [citation:2][citation:5]
    /// </summary>
    private void CopyCellStyle(ExcelRangeBase source, ExcelRangeBase target) {
        // 复制字体样式
        target.Style.Font.Size = source.Style.Font.Size;
        target.Style.Font.Bold = source.Style.Font.Bold;
        target.Style.Font.Italic = source.Style.Font.Italic;

        // 复制填充样式
        target.Style.Fill.PatternType = source.Style.Fill.PatternType;
        target.Style.Border = source.Style.Border;
        CopyBorderStyle(source.Style.Border, target.Style.Border);

        // 复制对齐方式
        target.Style.HorizontalAlignment = source.Style.HorizontalAlignment;
        target.Style.VerticalAlignment = source.Style.VerticalAlignment;
        target.Style.WrapText = source.Style.WrapText;

        // 复制数字格式
        target.Style.Numberformat.Format = source.Style.Numberformat.Format;
    }

    /// <summary>
    /// 复制边框样式 - 完整版本
    /// </summary>
    private void CopyBorderStyle(OfficeOpenXml.Style.Border sourceBorder, OfficeOpenXml.Style.Border targetBorder) {
        // 复制边框（样式+颜色）
        targetBorder.Left.Style = sourceBorder.Left.Style;
        if(sourceBorder.Left.Color.Theme!=null)
            targetBorder.Left.Color.SetColor(sourceBorder.Left.Color.Theme.Value);

        targetBorder.Right.Style = sourceBorder.Right.Style;
        if (sourceBorder.Right.Color.Theme != null)
            targetBorder.Right.Color.SetColor(sourceBorder.Right.Color.Theme.Value);

        targetBorder.Top.Style = sourceBorder.Top.Style;
        if (sourceBorder.Top.Color.Theme != null)
            targetBorder.Top.Color.SetColor(sourceBorder.Top.Color.Theme.Value);

        targetBorder.Bottom.Style = sourceBorder.Bottom.Style;
        if (sourceBorder.Bottom.Color.Theme != null)
            targetBorder.Bottom.Color.SetColor(sourceBorder.Bottom.Color.Theme.Value);
    }
    #endregion

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

    public void CopySheet(ExcelAttechFile[] excelAttechFiles, string targetSheetName, string targetStartCell) {

        var targetWorksheet = _excelPackage.Workbook.Worksheets[targetSheetName]
                            ?? _excelPackage.Workbook.Worksheets.Add(targetSheetName);

        var startCell = new ExcelCellAddress(targetStartCell);
        int currentCol = startCell.Column;

        foreach (var excelAttechFile in excelAttechFiles) {
            string sourcePath = excelAttechFile.FilePath;
            if (!File.Exists(sourcePath)) {
                throw new Exception($"文件不存在: {sourcePath}");
            }
            if (sourcePath.EndsWith(".xls")) {
                string newPath = ConvertXlsToXlsx(sourcePath);
                using (var sourcePackage = new ExcelPackage(new FileInfo(newPath))) {
                    var sourceWorksheet = sourcePackage.Workbook.Worksheets.FirstOrDefault();
                    if (sourceWorksheet?.Dimension == null) continue;

                    // 直接使用Copy方法复制整个工作表区域
                    var sourceRange = sourceWorksheet.Cells[sourceWorksheet.Dimension.Address];
                    var targetStart = targetWorksheet.Cells[startCell.Row, currentCol];

                    // 复制整个区域到目标位置
                    sourceRange.Copy(targetStart);

                    // 更新下一个位置
                    currentCol += sourceWorksheet.Dimension.Columns + 1;
                }
                File.Delete(newPath);
            }
            else {
                using (var sourcePackage = new ExcelPackage(new FileInfo(sourcePath))) {
                    var sourceWorksheet = sourcePackage.Workbook.Worksheets.FirstOrDefault();
                    if (sourceWorksheet?.Dimension == null) continue;

                    // 直接使用Copy方法复制整个工作表区域
                    var sourceRange = sourceWorksheet.Cells[sourceWorksheet.Dimension.Address];
                    var targetStart = targetWorksheet.Cells[startCell.Row, currentCol];

                    // 复制整个区域到目标位置
                    sourceRange.Copy(targetStart);

                    // 更新下一个位置
                    currentCol += sourceWorksheet.Dimension.Columns + 1;
                }
            }
        }
    }

    /// <summary>
    /// 将Xls格式文件转Excel（依赖本地Excel）
    /// </summary>
    public string ConvertXlsToXlsx(string inputPath) {
        object excelApp = null;
        object workbooks = null;
        object workbook = null;

        try {
            // 通过后期绑定创建Excel应用实例
            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            excelApp = Activator.CreateInstance(excelType);

            // 设置属性
            excelType.InvokeMember("Visible", BindingFlags.SetProperty, null, excelApp, new object[] { false });
            excelType.InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, excelApp, new object[] { false });

            // 获取Workbooks集合
            workbooks = excelType.InvokeMember("Workbooks", BindingFlags.GetProperty, null, excelApp, null);

            // 打开文件
            workbook = workbooks.GetType().InvokeMember("Open", BindingFlags.InvokeMethod, null, workbooks, new object[] { inputPath });

            // 生成新的XLSX文件名
            string outputPath = Path.Combine(
                Path.GetDirectoryName(inputPath),
                Guid.NewGuid().ToString() + ".xlsx"
            );

            // 保存为XLSX格式 (51代表xlsx格式)
            workbook.GetType().InvokeMember("SaveAs", BindingFlags.InvokeMethod, null, workbook, new object[] { outputPath, 51 });

            return outputPath;
        }
        catch (Exception ex) {
            Console.WriteLine($"转换失败: {ex.Message}");
            throw;
        }
        finally {
            // 清理COM对象
            if (workbook != null) {
                workbook.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, workbook, new object[] { false });
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            }
            if (workbooks != null) {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
            }
            if (excelApp != null) {
                excelApp.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, excelApp, null);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    /// <summary>
    /// 将Excel转PDF（依赖本地Excel）
    /// </summary>
    /// <param name="inputExcelPath">输入Excel路径（.xls/.xlsx）</param>
    /// <returns>输出PDF路径</returns>
    public string ConvertExcelToPdf(string inputExcelPath) {
        object excelApp = null;
        object workbooks = null;
        object workbook = null;

        try {
            // 验证输入文件
            if (!File.Exists(inputExcelPath))
                throw new FileNotFoundException("输入Excel文件不存在", inputExcelPath);

            // 1. 获取Excel应用类型（ProgID：Excel.Application）
            Type excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
                throw new Exception("未安装Excel或无法访问Excel COM组件");

            // 2. 创建Excel实例
            excelApp = Activator.CreateInstance(excelType);

            // 3. 设置属性：后台运行（不显示界面）、关闭警告
            excelType.InvokeMember("Visible", BindingFlags.SetProperty, null, excelApp, new object[] { false });
            excelType.InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, excelApp, new object[] { false });

            // 4. 获取Workbooks集合
            workbooks = excelType.InvokeMember("Workbooks", BindingFlags.GetProperty, null, excelApp, null);

            // 5. 打开Excel文件
            workbook = workbooks.GetType().InvokeMember(
                "Open",
                BindingFlags.InvokeMethod,
                null,
                workbooks,
                new object[] { inputExcelPath, Type.Missing, true } // 第三个参数：ReadOnly=true
            );

            // 6. 生成输出PDF路径（在原目录创建唯一文件名）
            string outputPdfPath = Path.Combine(
                Path.GetDirectoryName(inputExcelPath),
                $"{Guid.NewGuid()}.pdf"
            );

            // 7. 调用ExportAsFixedFormat导出PDF
            // 参数说明：
            // Type：0 = xlTypePDF（PDF格式）
            // Filename：输出路径
            // Quality：0 = xlQualityStandard（标准质量）
            workbook.GetType().InvokeMember(
                "ExportAsFixedFormat",
                BindingFlags.InvokeMethod,
                null,
                workbook,
                new object[] { 0, outputPdfPath, 0 } // Type=0(PDF), Filename, Quality=0(标准)
            );

            return outputPdfPath;
        }
        catch (Exception ex) {
            Console.WriteLine($"Excel转PDF失败：{ex.Message}");
            throw;
        }
        finally {
            // 强制清理COM对象，避免进程残留
            if (workbook != null) {
                // 关闭工作簿（不保存修改）
                workbook.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, workbook, new object[] { false });
                Marshal.ReleaseComObject(workbook);
            }
            if (workbooks != null) {
                Marshal.ReleaseComObject(workbooks);
            }
            if (excelApp != null) {
                // 退出Excel应用
                excelApp.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, excelApp, null);
                Marshal.ReleaseComObject(excelApp);
            }

            // 强制垃圾回收
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    /// <summary>
    /// 将Word转PDF（依赖本地Word）
    /// </summary>
    /// <param name="inputWordPath">输入Word路径（.doc/.docx）</param>
    /// <returns>输出PDF路径</returns>
    public string ConvertWordToPdf(string inputWordPath) {
        object wordApp = null;
        object documents = null;
        object document = null;

        try {
            // 验证输入文件
            if (!File.Exists(inputWordPath))
                throw new FileNotFoundException("输入Word文件不存在", inputWordPath);

            // 1. 获取Word应用类型（ProgID：Word.Application）
            Type wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null)
                throw new Exception("未安装Word或无法访问Word COM组件");

            // 2. 创建Word实例
            wordApp = Activator.CreateInstance(wordType);

            // 3. 设置属性：后台运行、关闭警告
            wordType.InvokeMember("Visible", BindingFlags.SetProperty, null, wordApp, new object[] { false });
            wordType.InvokeMember("DisplayAlerts", BindingFlags.SetProperty, null, wordApp, new object[] { 0 }); // 0=wdAlertsNone

            // 4. 获取Documents集合
            documents = wordType.InvokeMember("Documents", BindingFlags.GetProperty, null, wordApp, null);

            // 5. 打开Word文件（只读模式）
            document = documents.GetType().InvokeMember(
                "Open",
                BindingFlags.InvokeMethod,
                null,
                documents,
                new object[] { inputWordPath, Type.Missing, true } // 第三个参数：ReadOnly=true
            );

            // 6. 生成输出PDF路径
            string outputPdfPath = Path.Combine(
                Path.GetDirectoryName(inputWordPath),
                $"{Guid.NewGuid()}.pdf"
            );

            // 7. 调用SaveAs2保存为PDF
            // 参数说明：
            // FileName：输出路径
            // FileFormat：17 = wdFormatPDF（PDF格式）
            document.GetType().InvokeMember(
                "SaveAs2",
                BindingFlags.InvokeMethod,
                null,
                document,
                new object[] { outputPdfPath, 17 } // FileName, FileFormat=17(PDF)
            );

            return outputPdfPath;
        }
        catch (Exception ex) {
            Console.WriteLine($"Word转PDF失败：{ex.Message}");
            throw;
        }
        finally {
            // 清理COM对象
            if (document != null) {
                // 关闭文档（不保存修改）
                document.GetType().InvokeMember("Close", BindingFlags.InvokeMethod, null, document, new object[] { 0 }); // 0=wdDoNotSaveChanges
                Marshal.ReleaseComObject(document);
            }
            if (documents != null) {
                Marshal.ReleaseComObject(documents);
            }
            if (wordApp != null) {
                // 退出Word应用
                wordApp.GetType().InvokeMember("Quit", BindingFlags.InvokeMethod, null, wordApp, null);
                Marshal.ReleaseComObject(wordApp);
            }

            // 强制垃圾回收
            GC.Collect();
            GC.WaitForPendingFinalizers();
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


    /// <summary>
    /// 替换文件中所有公式为所见值，保持单元格式不变
    /// </summary>
    public void ReplaceFormula() {
        // 遍历所有工作表
        foreach (var worksheet in _excelPackage.Workbook.Worksheets) {
            // 跳过空工作表
            if (worksheet.Dimension == null) continue;

            // 获取数据区域范围
            int startRow = worksheet.Dimension.Start.Row;
            int endRow = worksheet.Dimension.End.Row;
            int startCol = worksheet.Dimension.Start.Column;
            int endCol = worksheet.Dimension.End.Column;
            worksheet.Calculate();

            for (int row = startRow; row <= endRow; row++) {
                for (int col = startCol; col <= endCol; col++) {
                    var cell = worksheet.Cells[row, col];
                    // 检查单元格是否包含公式
                    if (!string.IsNullOrEmpty(cell.Formula)) {
                        var cellValue = cell.Value;
                        // 清除公式
                        cell.Formula = null;
                        // 恢复值
                        cell.Value = cellValue;
                    }
                }
            }
        }
    }

    /// <summary>
    /// 向右复制N次（支持多列区域）
    /// </summary>
    /// <param name="sheetName">Sheet名称</param>
    /// <param name="cellsZone">单元格区域（如"A1:C3"）</param>
    /// <param name="copyCount">复制次数</param>
    public void CopyColumnsCells(string sheetName, string cellsZone, int copyCount) {
        var targetWorksheet = _excelPackage.Workbook.Worksheets[sheetName]
                            ?? _excelPackage.Workbook.Worksheets.Add(sheetName);
        // 获取原始单元格区域
        var originalRange = targetWorksheet.Cells[cellsZone];

        // 验证：复制次数必须为正数，且原始区域有效
        if (copyCount <= 0 || originalRange == null) {
            return;
        }

        // 获取原始区域的行列信息
        int originalStartCol = originalRange.Start.Column;   // 原始区域起始列
        int originalEndCol = originalRange.End.Column;       // 原始区域结束列
        int originalColumnCount = originalEndCol - originalStartCol + 1;  // 原始区域列数
        int startRow = originalRange.Start.Row;              // 原始区域起始行
        int endRow = originalRange.End.Row;                  // 原始区域结束行

        // 循环复制到右侧（共复制copyCount次）
        for (int i = 1; i <= copyCount; i++) {
            // 计算当前复制块的目标列范围
            int targetStartCol = originalEndCol + 1 + (i - 1) * originalColumnCount;  // 目标起始列
            int targetEndCol = targetStartCol + originalColumnCount - 1;              // 目标结束列（与原始区域列数一致）

            // 定义目标区域（与原始区域行列范围相同，仅列整体右移）
            var targetRange = targetWorksheet.Cells[startRow, targetStartCol, endRow, targetEndCol];
            // 复制原始区域到目标区域（包含值、公式、格式等）
            originalRange.Copy(targetRange);
        }
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

    private (ExcelDrawing[] embedObjects, bool isImage, int width) CreateEmbedObject(
        ExcelWorksheet worksheet,
        ExcelAttechFile excelAttechFile, double targetRowHeight, bool attMode = false) {
        string fileName = excelAttechFile.FileName;
        string filePath = excelAttechFile.FilePath;

        var extension = Path.GetExtension(filePath).ToLower();
        var imageTypes = new[] { ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".pdf" };
        int width = (int)(targetRowHeight);
        //word
        if (new string[] { ".doc", ".docx" }.Contains(extension)) {
            if (attMode) {
                using (var fileStream = new FileStream(filePath, FileMode.Open)) {
                    var oleObject = worksheet.Drawings.AddOleObject(
                        Guid.NewGuid().ToString(),
                        fileStream,
                        fileName,
                        parameters => {
                            // 在这里设置参数
                            parameters.DisplayAsIcon = false; // 开启显示图标模式
                            parameters.ProgId = "Word.Document"; // 添加这两行
                        }
                    );
                    oleObject.SetSize(width, width);
                    return (new ExcelDrawing[] { oleObject }, false, width);
                }
            }
            else {
                //1.转pdf
                string pdfPath = ConvertWordToPdf(filePath);
                var listStream = FileToImageHelper.ConvertAllPagesToImageStreams(pdfPath);
                List<ExcelPicture> list = new List<ExcelPicture>();
                foreach (var stream in listStream) {
                    using (stream) {
                        var picture = worksheet.Drawings.AddPicture(
                          Guid.NewGuid().ToString(),
                          stream
                        );
                        if (imageTypes.Contains(extension)) {
                            // 添加图片后立即缩放
                            width = ScaleImageToCell(picture, targetRowHeight);
                        }
                        list.Add(picture);
                    }
                }
                File.Delete(pdfPath);
                return (list.ToArray(), true, width);
            }
        }
        //excel
        else if (new string[] { ".xls", ".xlsx" }.Contains(extension)) {
            if (attMode) {
                using (var fileStream = new FileStream(filePath, FileMode.Open)) {
                    var oleObject = worksheet.Drawings.AddOleObject(
                        Guid.NewGuid().ToString(),
                        fileStream,
                        fileName,
                    parameters => {
                        // 在这里设置参数
                        parameters.DisplayAsIcon = true; // 开启显示图标模式
                        parameters.ProgId = "Excel.Sheet";
                    }
                    );
                    oleObject.SetSize(width, width);
                    oleObject.Locked = true;
                    return (new ExcelDrawing[] { oleObject }, false, width);
                }
            }
            else {
                //1.转pdf
                string pdfPath = ConvertExcelToPdf(filePath);
                var listStream = FileToImageHelper.ConvertAllPagesToImageStreams(pdfPath);
                List<ExcelPicture> list = new List<ExcelPicture>();
                foreach (var stream in listStream) {
                    using (stream) {
                        var picture = worksheet.Drawings.AddPicture(
                          Guid.NewGuid().ToString(),
                          stream
                        );
                        if (imageTypes.Contains(extension)) {
                            // 添加图片后立即缩放
                            width = ScaleImageToCell(picture, targetRowHeight);
                        }
                        list.Add(picture);
                    }
                }
                File.Delete(pdfPath);
                return (list.ToArray(), true, width);
            }
        }
        else if (imageTypes.Contains(extension)) {
            //pdf
            if (extension.Contains(".pdf")) {
                if (attMode) {
                    using (var fileStream = new FileStream(filePath, FileMode.Open)) {
                        var oleObject = worksheet.Drawings.AddOleObject(
                        Guid.NewGuid().ToString(),
                        fileStream, fileName,
                        parameters => {
                            // 在这里设置参数
                            parameters.DisplayAsIcon = true; // 开启显示图标模式
                            parameters.ProgId = "AcroExch.Document";
                        }
                    );
                        oleObject.SetSize(width, width);
                        oleObject.Locked = true;
                        return (new ExcelDrawing[] { oleObject }, false, width);
                    }
                }
                else {
                    //using (var stream = FileToImageHelper.ConvertFirstPageToImageStream(filePath)) {
                    //    var picture = worksheet.Drawings.AddPicture(
                    //      Guid.NewGuid().ToString(),
                    //      stream
                    //    );
                    //    if (imageTypes.Contains(extension)) {
                    //        // 添加图片后立即缩放
                    //        width = ScaleImageToCell(picture, targetRowHeight);
                    //    }
                    //    return (new ExcelDrawing[] { picture }, true, width);
                    //}
                    var listStream = FileToImageHelper.ConvertAllPagesToImageStreams(filePath);
                    List<ExcelPicture> list = new List<ExcelPicture>();
                    foreach (var stream in listStream) {
                        using (stream) {
                            var picture = worksheet.Drawings.AddPicture(
                              Guid.NewGuid().ToString(),
                              stream
                            );
                            if (imageTypes.Contains(extension)) {
                                // 添加图片后立即缩放
                                width = ScaleImageToCell(picture, targetRowHeight);
                            }
                            list.Add(picture);
                        }
                    }
                    return (list.ToArray(), true, width);
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
                    return (new ExcelDrawing[] { picture }, true, width);
                }
            }
        }
        else {
            using (var fileStream = new FileStream(filePath, FileMode.Open)) {
                var oleObject = worksheet.Drawings.AddOleObject(
                    Guid.NewGuid().ToString(),
                    fileStream, fileName,
                    parameters => {
                        // 在这里设置参数
                        parameters.DisplayAsIcon = true; // 开启显示图标模式
                        parameters.ProgId = "Package";
                    }
                );
                oleObject.SetSize(width, width);
                oleObject.Locked = true;
                return (new ExcelDrawing[] { oleObject }, false, width);
            }
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

public class ExcelAttechFile {
    public string FileName {  get; set; }
    public string FilePath { get; set; }
}