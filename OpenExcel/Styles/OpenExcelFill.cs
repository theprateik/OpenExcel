using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel.Styles
{
    public class OpenExcelFill
    {
        public OpenExcelFill(string uid)
        {
            UID = uid;
        }

        public string UID { get; }

        public EnumValue<PatternValues> PatternType { get; set; }

        public ForegroundColor ForegroundColor { get; set; }

        public BackgroundColor BackgroundColor { get; set; }
    }
}
