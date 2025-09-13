using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf;
using System;
using System.Reflection.PortableExecutable;
using Microsoft.AspNetCore.Components.Forms;
using iText.Forms.Form.Element;
using Aspose.Pdf.Text;
using iText.Kernel.Pdf.Annot;

namespace Meiam.System.Interfaces
{
    internal class ITextHelper
    {
        public string ExtractAllText(Aspose.Pdf.Document pdfDoc)
        {
            // 创建一个文本吸收器（抓取所有文本）
            TextAbsorber textAbsorber = new TextAbsorber();

            // 让吸收器访问所有页面
            pdfDoc.Pages.Accept(textAbsorber);

            // 返回提取的文本
            return textAbsorber.Text;
        }
        public string ExtractTextRightOfKeyword(string allText, string keyword)
        {
            if (string.IsNullOrEmpty(allText) || string.IsNullOrEmpty(keyword))
                return "文本或关键字为空";

            // 按行拆分
            var lines = allText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);

            foreach (var line in lines)
            {
                int index = line.IndexOf(keyword);
                if (index != -1)
                {
                    // 提取该行中关键词右边的文本
                    int startIndex = index + keyword.Length;
                    if (startIndex >= line.Length)
                        return "关键字后无内容";

                    return line.Substring(startIndex).Trim();
                }
            }

            return $"未找到关键字：{keyword}";
        }

        public void ReplacePdfText(string searchText, string replaceText, Aspose.Pdf.Document pdfDoc)
        {
            // 创建文本搜索器
            TextFragmentAbsorber absorber = new TextFragmentAbsorber(searchText);
            pdfDoc.Pages.Accept(absorber);

            // 遍历找到的文本片段并替换
            foreach (TextFragment fragment in absorber.TextFragments)
            {
                // 读取原来的字体
                Aspose.Pdf.Text.Font originalFont = FontRepository.FindFont("SimSun");// fragment.TextState.Font;
                float originalSize = fragment.TextState.FontSize;
                Aspose.Pdf.Color originalColor = fragment.TextState.ForegroundColor;
                // 替换文本
                fragment.Text = replaceText;

                // 保持原字体、大小、颜色
                fragment.TextState.Font = originalFont;
                fragment.TextState.FontSize = originalSize;
                fragment.TextState.ForegroundColor = originalColor;
            }

        }
        public static void RemoveWatermark(string inputPdfPath, string outputPdfPath)
        {
            using (var reader = new PdfReader(inputPdfPath))
            using (var writer = new PdfWriter(outputPdfPath))
            using (var pdfDoc = new PdfDocument(reader, writer))
            {
                // 设定白色矩形的位置和大小
                //float x = 0;          // 左上角横坐标
                //float y = pdfDoc.GetDefaultPageSize().GetTop();  // 页面顶部纵坐标
                float width = 400;    // 矩形宽度
                float height = 20; // 按视觉高度设置

                // 遍历每一页
                for (int pageNum = 1; pageNum <= pdfDoc.GetNumberOfPages(); pageNum++)
                {
                    var page = pdfDoc.GetPage(pageNum);
                    var cropBox = page.GetCropBox();
                    float visibleTop = cropBox.GetTop(); // 可见区域的顶部y坐标（而非原始高度600）
                    float x = cropBox.GetLeft(); // 可见区域的左边界
                    var canvas = new PdfCanvas(page);
                    // 设置矩形的填充颜色为白色，并绘制矩形
                    canvas.SaveState()
                          .SetFillColorRgb(1f, 1f, 1f) // 白色
                          .Rectangle(x, visibleTop - height, width, height) // 左上角矩形
                          .Fill()
                          .RestoreState();
                }
            }
        }
    }
}
