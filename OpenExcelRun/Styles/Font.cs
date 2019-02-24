using OpenExcel.Styles;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcelRun.Styles
{
    public static class Font
    {
        public static OpenExcelFont F1 = new OpenExcelFont("8379333F-2626-48FD-8464-3CD68D508BAF") { FontSize = 12, Color = "Red", Bold = true };

        public static OpenExcelFont F2 = new OpenExcelFont("12467A82-2281-4FBF-BBA5-45F230BC5E43") { FontSize = 8, Color = "Red", Italic = true };

    }
}
