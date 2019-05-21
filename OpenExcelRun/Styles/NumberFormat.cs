using OpenExcel.Styles;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcelRun.Styles
{
    public static class NumberFormat
    {
        public static OpenExcelNumberingFormat N1 = new OpenExcelNumberingFormat(Guid.NewGuid().ToString()) { FormatCode = "0.000" };

        public static OpenExcelNumberingFormat N2 = new OpenExcelNumberingFormat(Guid.NewGuid().ToString()) { FormatCode = "[$-409]d\\-mmm\\-yyyy;@" };

        public static OpenExcelNumberingFormat N3 = new OpenExcelNumberingFormat(Guid.NewGuid().ToString()) { FormatCode = "0.00000" };
    }
}
