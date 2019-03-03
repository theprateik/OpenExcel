using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcel.Styles;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcelRun.Styles
{
    public static class Fill
    {
        public static OpenExcelFill F1 = new OpenExcelFill(Guid.NewGuid().ToString()) { PatternType = PatternValues.Solid, ForegroundColor = new ForegroundColor() { Rgb = "808080" } };

        public static OpenExcelFill F2 = new OpenExcelFill(Guid.NewGuid().ToString()) { PatternType = PatternValues.Solid, ForegroundColor = new ForegroundColor() { Rgb = "f90000" } };
    }
}
