using ClosedXML.Excel;
using System.Drawing;
using System.Threading.Tasks;

namespace ExcelYourself.Core
{
    public interface IImageToExcelConverter
    {
        XLWorkbook Convert(Bitmap bitmap);
    }
}