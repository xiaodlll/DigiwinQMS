using DocumentFormat.OpenXml.Office2010.ExcelAc;
using Meiam.System.Common;
using NPOI.HPSF;
using NPOI.Util;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;
using System;
using System.IO;
using System.Linq;

public class ExcelHelper : IDisposable {
    private readonly ExcelPackage _excelPackage;
    private readonly FileInfo _fileInfo;
    private const int BaseRowHeight = 15; // 默认行高（磅）
    private const int BaseColumnWidth = 10; // 默认列宽（字符数）

    public ExcelHelper(string file) {
        _fileInfo = new FileInfo(file);
        _excelPackage = _fileInfo.Exists ?
            new ExcelPackage(_fileInfo) :
            new ExcelPackage();
        ExcelPackage.License.SetNonCommercialPersonal("xiaodl");
    }

    public void AddTextToCell(string sheetName, string cellAddress, string text) {
        var worksheet = GetOrCreateWorksheet(sheetName);
        var cell = worksheet.Cells[cellAddress];

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

        double startLeft = 5;

        // 设置基础样式
        cell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
        worksheet.Row(startRow).CustomHeight = true;

        foreach (var filePath in fullFilePaths.Where(File.Exists)) {
            try {
                var (embedObject, isImage, width) = CreateEmbedObject(worksheet, filePath, worksheet.Row(startRow).Height);

                // 设置对象位置
                embedObject.SetPosition(
                    startRow - 1,
                    5,
                    startCol - 1,
                    (int)(startLeft)
                );
                startLeft += width + 5;
            }
            catch (Exception ex) {
                AddErrorComment(cell, $"文件嵌入失败: {ex.Message}");
            }
        }
    }
    public void InsertRows(string sheetName, int rowCount, int targetRow) {
        // 获取指定工作表
        var worksheet = GetOrCreateWorksheet(sheetName);
        // 在目标位置插入新行
        worksheet.InsertRow(targetRow, rowCount);
    }

    public void CopyRows(string sheetName, int[] sourceRow, int targetRow) {
        // 获取指定工作表
        var worksheet = GetOrCreateWorksheet(sheetName);
        int cols = worksheet.Dimension.End.Column; // 获取总列数

        // 将源行按升序排序以便正确处理行号偏移
        var sortedSourceRows = sourceRow.OrderBy(r => r).ToArray();

        int rowOffset = 0; // 跟踪因插入导致的行号偏移
        int currentTargetRow = targetRow; // 当前插入的目标行

        foreach (int originalRow in sortedSourceRows) {
            // 计算调整后的源行号（考虑之前的插入操作）
            int adjustedRow = originalRow + rowOffset;

            // 在目标位置插入新行
            worksheet.InsertRow(currentTargetRow, 1);

            // 复制源行内容到目标行
            var sourceRange = worksheet.Cells[adjustedRow, 1, adjustedRow, cols];
            var targetRange = worksheet.Cells[currentTargetRow, 1];
            sourceRange.Copy(targetRange);

            // 如果插入位置在源行之前或同一行，则后续源行号需调整
            if (currentTargetRow <= adjustedRow) {
                rowOffset++;
            }

            currentTargetRow++; // 更新下一个目标行位置
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
        var imageTypes = new[] { ".png", ".jpg", ".jpeg", ".gif", ".bmp" };
        int width = (int)(targetRowHeight);
        if (imageTypes.Contains(extension)) {
            var picture = worksheet.Drawings.AddPicture(
                Guid.NewGuid().ToString(),
                new FileInfo(filePath)
            ); 
            if (imageTypes.Contains(extension)) {
                // 添加图片后立即缩放
                width = ScaleImageToCell(picture, targetRowHeight);
            }
            return (picture, true, width);
        }
        else {
            var oleObject = worksheet.Drawings.AddOleObject(
                Guid.NewGuid().ToString(),
                new FileInfo(filePath)
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

    private void AddErrorComment(ExcelRange cell, string message) {
        var comment = cell.AddComment(message, "System");
        comment.AutoFit = true;
        comment.BackgroundColor = System.Drawing.Color.LightYellow;
    }
}