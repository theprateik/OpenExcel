using OpenExcel.Styles;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcelRun.Styles
{
    public static class Font
    {
        public static OpenExcelFont F1 = new OpenExcelFont(Guid.NewGuid().ToString()) { FontName = "Calibri", FontSize = 12, Color = "f90000", Bold = true };

        public static OpenExcelFont F2 = new OpenExcelFont(Guid.NewGuid().ToString()) { FontSize = 8, Color = "f90000", Italic = true };
    }
}
