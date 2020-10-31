using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelYourself.Web.Models
{
    public class FileUploadModel
    {
        [Required(ErrorMessage = "You need to select an image to convert.")]
        [Display(Name = "Choose image to convert")]
        public IFormFile FormFile { get; set; }
    }
}
