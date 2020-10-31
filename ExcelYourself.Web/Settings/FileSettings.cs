using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelYourself.Web.Settings
{
    public class FileSettings
    {
        public long FileSizeLimit { get; set; } = 8 * 1024 * 1024;
        public int DesiredWidth { get; set; } = 420;
        public int DesiredHeight { get; set; } = 188;
    }
}
