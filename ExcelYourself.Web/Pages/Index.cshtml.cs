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

        public async Task<IActionResult> OnPostUploadAsync()
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
                using (var memoryStream = new MemoryStream())
                {
                    await FileUpload.FormFile.CopyToAsync(memoryStream);
                    using var image = SKBitmap.Decode(memoryStream);

                    using (var workbook = _converter.Convert(ResizeImage(image)))
                    {
                        return workbook.Deliver($"{Path.GetFileNameWithoutExtension(FileUpload.FormFile.FileName)}.xlsx");
                    }
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "{ErrorName} error encountered while converting image.", ex.GetType().Name);
                ModelState.AddModelError(FileUpload.FormFile.Name, $"Oops, something unexpected happened and I couldn't convert you image.");
                return Page();
            }
        }

        private Bitmap ResizeImage(Image source)
        {
            // 421 x (568 / 3) = 421 x 189
            float scale = Math.Min((float)_options.Value.DesiredWidth / source.Width, (float)_options.Value.DesiredHeight / source.Height);

            var bitmap = new Bitmap(
                (int)Math.Min(_options.Value.DesiredWidth, scale * source.Width),
                (int)Math.Min(_options.Value.DesiredHeight, scale * source.Height));
            bitmap.SetResolution(source.HorizontalResolution, source.VerticalResolution);

            var destRect = new Rectangle(0, 0, bitmap.Width, bitmap.Height);

            using (var graphic = Graphics.FromImage(bitmap))
            {
                graphic.CompositingMode = CompositingMode.SourceCopy;
                graphic.CompositingQuality = CompositingQuality.HighQuality;
                graphic.InterpolationMode = InterpolationMode.HighQualityBicubic;
                graphic.SmoothingMode = SmoothingMode.HighQuality;
                graphic.PixelOffsetMode = PixelOffsetMode.HighQuality;

                using (var wrapMode = new ImageAttributes())
                {
                    graphic.DrawImage(source, destRect, 0, 0, source.Width, source.Height, GraphicsUnit.Pixel, wrapMode);
                }
            }
            return bitmap;
        }
    }
}
