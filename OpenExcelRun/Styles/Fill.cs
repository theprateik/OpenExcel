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
        public static OpenExcelFill F1 => new OpenExcelFill("{A7DAE395-4BC1-4FE8-904C-1D8686950892}") { PatternType = PatternValues.Solid, ForegroundColor = new ForegroundColor() { Rgb = "808080" } };

        public static OpenExcelFill F2 => new OpenExcelFill("{759649CB-325C-4AE0-86AB-CB242922EAF9}") { PatternType = PatternValues.Solid, ForegroundColor = new ForegroundColor() { Rgb = "f90000" } };
    }
}
