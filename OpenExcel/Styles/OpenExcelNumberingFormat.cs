using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenExcel.Styles
{
    public class OpenExcelNumberingFormat
    {
        public OpenExcelNumberingFormat(string uid)
        {
            UID = uid;
        }
        public string UID { get; }
        public string FormatCode { get; set; }
    }
}
