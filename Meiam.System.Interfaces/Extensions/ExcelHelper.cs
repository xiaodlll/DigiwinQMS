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
        worksheet.Cells[cellAddress].Value = text;
    }

    public void AddAttachsToCell(string sheetName, string cellAddress, string[] filePaths) {
        var worksheet = GetOrCreateWorksheet(sheetName);
        var cell = worksheet.Cells[cellAddress];
        var startRow = cell.Start.Row;
        var startCol = cell.Start.Column;

        double startLeft = 5;

        // 设置基础样式
        cell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
        worksheet.Row(startRow).CustomHeight = true;

        foreach (var filePath in filePaths.Where(File.Exists)) {
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

    public void CopyRow(string sheetName, int[] sourceRow, int targetRow) {
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