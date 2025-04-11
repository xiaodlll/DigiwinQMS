using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Pdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Meiam.System.Interfaces.Extensions
{
    public class ITextHelper
    {
        public static void RemoveWatermark(string inputPdfPath, string outputPdfPath)
        {
            using (var reader = new PdfReader(inputPdfPath))
            using (var writer = new PdfWriter(outputPdfPath))
            using (var pdfDoc = new PdfDocument(reader, writer))
            {
                // 设定白色矩形的位置和大小
                float x = 0;          // 左上角横坐标
                float y = pdfDoc.GetDefaultPageSize().GetTop();  // 页面顶部纵坐标
                float width = 400;    // 矩形宽度
                float height = 265;   // 矩形高度

                // 遍历每一页
                for (int pageNum = 1; pageNum <= pdfDoc.GetNumberOfPages(); pageNum++)
                {
                    var page = pdfDoc.GetPage(pageNum);
                    var canvas = new PdfCanvas(page);

                    // 设置矩形的填充颜色为白色，并绘制矩形
                    canvas.SaveState()
                          .SetFillColorRgb(1f, 1f, 1f) // 白色
                          .Rectangle(x, y - height, width, height) // 左上角矩形
                          .Fill()
                          .RestoreState();
                }
            }
        }
    }
}
