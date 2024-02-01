using ClosedXML.Excel;
using ClosedXML.Extensions;

using ExcelYourself.Core;
using ExcelYourself.Web.Models;
using ExcelYourself.Web.Settings;

using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;

using SkiaSharp;

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Net;
using System.Threading.Tasks;

namespace ExcelYourself.Web.Pages
{
    public class IndexModel : PageModel
    {
        private readonly IOptions<FileSettings> _options;
        private readonly IImageToExcelConverter _converter;
        private readonly ILogger<IndexModel> _logger;

        [BindProperty]
        public FileUploadModel FileUpload { get; set; }

        public IndexModel(IImageToExcelConverter converter, IOptions<FileSettings> options, ILogger<IndexModel> logger)
        {
            _converter = converter;
            _options = options;
            _logger = logger;
        }

        public void OnGet()
        {

        }

        public IActionResult OnPostUpload()
        {
            if (!ModelState.IsValid)
            {
                return Page();
            }

            var safeFileName = WebUtility.HtmlEncode(FileUpload.FormFile.FileName);

            if (FileUpload.FormFile.Length == 0)
            {
                ModelState.AddModelError(FileUpload.FormFile.Name, $"File {safeFileName} is empty.");
            }

            if (FileUpload.FormFile.Length > _options.Value.FileSizeLimit)
            {
                var megabyteSizeLimit = _options.Value.FileSizeLimit / 1024 / 1024;
                ModelState.AddModelError(FileUpload.FormFile.Name, $"File {safeFileName} exceeds {megabyteSizeLimit:N1} MB.");
            }

            if (!ModelState.IsValid)
            {
                return Page();
            }

            try
            {
                //using var memoryStream = new MemoryStream();
                //await FileUpload.FormFile.CopyToAsync(memoryStream);
                //memoryStream.Seek(0, SeekOrigin.Begin);
                var image = SKBitmap.Decode(FileUpload.FormFile.OpenReadStream());

                using XLWorkbook workbook = _converter.Convert(ResizeImage(image));
                return workbook.Deliver($"{Path.ChangeExtension(FileUpload.FormFile.FileName, "xlsx")}");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "{ErrorName} error encountered while converting image.", ex.GetType().Name);
                ModelState.AddModelError(FileUpload.FormFile.Name, $"Oops, something unexpected happened and I couldn't convert your image.");
                return Page();
            }
        }

        private SKBitmap ResizeImage(SKBitmap source)
        {
            // 421 x (568 / 3) = 421 x 189
            float scale = Math.Min((float)_options.Value.DesiredWidth / source.Width, (float)_options.Value.DesiredHeight / source.Height);

            var targetImageInfo = new SKImageInfo(
                (int)Math.Min(_options.Value.DesiredWidth, scale * source.Width),
                (int)Math.Min(_options.Value.DesiredHeight, scale * source.Height));

            return source.Resize(targetImageInfo, SKFilterQuality.High);
        }
    }
}
