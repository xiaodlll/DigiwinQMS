using OfficeOpenXml;
using OfficeOpenXml.Drawing;
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

        // 初始化尺寸参数
        var totalHeight = 0.0;
        var maxWidth = 0.0;

        // 设置基础样式
        cell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Top;
        worksheet.Row(startRow).CustomHeight = true;

        foreach (var filePath in filePaths.Where(File.Exists)) {
            try {
                var (embedObject, isImage) = CreateEmbedObject(worksheet, filePath);

                // 计算对象显示尺寸（保持比例）
                var width = 30;
                var height = 30;

                // 调整对象尺寸以适应列宽
                if (width > 200) {
                    var ratio = 200 / width;
                    width *= ratio;
                    height *= ratio;
                }

                // 设置对象位置（垂直排列）
                embedObject.SetPosition(
                    startRow - 1,
                    (int)(totalHeight * BaseRowHeight),
                    startCol - 1,
                    0
                );

                // 更新尺寸跟踪
                totalHeight += height / BaseRowHeight;
                maxWidth = Math.Max(maxWidth, width);

                // 自动调整行高
                worksheet.Row(startRow).Height =
                    Math.Max(worksheet.Row(startRow).Height,
                    (double)(totalHeight * BaseRowHeight));
            }
            catch (Exception ex) {
                AddErrorComment(cell, $"文件嵌入失败: {ex.Message}");
            }
        }

        // 自动调整列宽
        var newWidth = Math.Max(
            worksheet.Column(startCol).Width,
            (double)(maxWidth / 7.5)  // 转换为Excel列宽单位
        );
        worksheet.Column(startCol).Width = newWidth;
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

    private (ExcelDrawing embedObject, bool isImage) CreateEmbedObject(
        ExcelWorksheet worksheet,
        string filePath) {
        var extension = Path.GetExtension(filePath).ToLower();
        var imageTypes = new[] { ".png", ".jpg", ".jpeg", ".gif", ".bmp" };

        if (imageTypes.Contains(extension)) {
            var picture = worksheet.Drawings.AddPicture(
                Guid.NewGuid().ToString(),
                new FileInfo(filePath)
            );
            return (picture, true);
        }
        else {
            var oleObject = worksheet.Drawings.AddOleObject(
                Guid.NewGuid().ToString(),
                new FileInfo(filePath)
            );
            return (oleObject, false);
        }
    }

    private void AddErrorComment(ExcelRange cell, string message) {
        var comment = cell.AddComment(message, "System");
        comment.AutoFit = true;
        comment.BackgroundColor = System.Drawing.Color.LightYellow;
    }
}