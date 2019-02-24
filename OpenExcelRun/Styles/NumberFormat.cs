using OpenExcel.Styles;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcelRun.Styles
{
    public static class NumberFormat
    {
        public static OpenExcelNumberingFormat N1 = new OpenExcelNumberingFormat(Guid.NewGuid().ToString()) { FormatCode = "0.000" };
    }
}
