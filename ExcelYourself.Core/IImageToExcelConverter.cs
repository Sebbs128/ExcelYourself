using ClosedXML.Excel;

using SkiaSharp;

namespace ExcelYourself.Core
{
    public interface IImageToExcelConverter
    {
        XLWorkbook Convert(SKBitmap bitmap);
    }
}