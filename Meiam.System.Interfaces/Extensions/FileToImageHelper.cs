using PdfiumViewer;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

public class FileToImageHelper {

    /// <summary>
    /// 将 PDF 的第 1 页渲染为 PNG 图片，并返回为 MemoryStream。
    /// </summary>
    /// <param name="pdfPath">PDF 文件路径</param>
    /// <param name="dpi">渲染分辨率（DPI）</param>
    /// <returns>包含 PNG 图片的 MemoryStream</returns>
    public static MemoryStream ConvertFirstPageToImageStream(string pdfPath, int dpi = 300) {
        // 加载 PDF 文档
        using var document = PdfDocument.Load(pdfPath);

        // 渲染第一页（页码从0开始）
        using var image = document.Render(0, dpi, dpi, PdfRenderFlags.CorrectFromDpi);

        // 写入图片到内存流
        var stream = new MemoryStream();
        image.Save(stream, ImageFormat.Png);
        stream.Position = 0; // 重置位置
        return stream;
    }

    /// <summary>
    /// 将 PDF 的所有页渲染为 PNG 图片，并返回包含所有页图片的 MemoryStream 列表
    /// </summary>
    /// <param name="pdfPath">PDF 文件路径</param>
    /// <param name="dpi">渲染分辨率（DPI）</param>
    /// <returns>包含所有页 PNG 图片的 MemoryStream 列表</returns>
    public static List<MemoryStream> ConvertAllPagesToImageStreams(string pdfPath, int dpi = 300) {
        var imageStreams = new List<MemoryStream>();

        using var document = PdfDocument.Load(pdfPath);

        for (int pageIndex = 0; pageIndex < document.PageCount; pageIndex++) {
           
            using var image = document.Render(pageIndex, dpi, dpi, PdfRenderFlags.CorrectFromDpi);
            var stream = new MemoryStream();
            image.Save(stream, ImageFormat.Png);
            stream.Position = 0;
            imageStreams.Add(stream);
        }

        return imageStreams;
    }
}