using ClosedXML.Excel;

using SkiaSharp;

using System.Collections.Generic;
using System.Linq;

namespace ExcelYourself.Core
{
    public class ImageToExcelConverter : IImageToExcelConverter
    {
        public ImageToExcelConverter()
        {
        }

        public XLWorkbook Convert(SKBitmap bitmap)
        {
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Sheet 1");

            for (int x = 0; x < bitmap.Width; x++)
            {
                string col = ToBase26(x + 1);
                for (int y = 0; y < bitmap.Height; y++)
                {
                    var pixel = bitmap.GetPixel(x, y);

                    int rowRed = y * 3 + 1;
                    int rowGreen = y * 3 + 2;
                    int rowBlue = y * 3 + 3;

                    worksheet.Cell($"{col}{rowRed}").Value = pixel.Red;
                    worksheet.Cell($"{col}{rowRed}").Style.Fill.BackgroundColor = XLColor.FromArgb(pixel.Red, 0, 0);
                    worksheet.Cell($"{col}{rowGreen}").Value = pixel.Green;
                    worksheet.Cell($"{col}{rowGreen}").Style.Fill.BackgroundColor = XLColor.FromArgb(0, pixel.Green, 0);
                    worksheet.Cell($"{col}{rowBlue}").Value = pixel.Blue;
                    worksheet.Cell($"{col}{rowBlue}").Style.Fill.BackgroundColor = XLColor.FromArgb(0, 0, pixel.Blue);
                }
            }

            return workbook;
        }

        private string ToBase26(int number)
        {
            var array = new LinkedList<int>();

            while (number > 26)
            {
                int value = number % 26;
                if (value == 0)
                {
                    number = number / 26 - 1;
                    array.AddFirst(26);
                }
                else
                {
                    number /= 26;
                    array.AddFirst(value);
                }
            }

            if (number > 0)
            {
                array.AddFirst(number);
            }
            return new string(array.Select(s => (char)('A' + s - 1)).ToArray());
        }
    }
}
