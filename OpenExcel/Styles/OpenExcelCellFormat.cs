using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel.Styles
{
    public class OpenExcelCellFormat
    {
        public OpenExcelCellFormat(string uid)
        {
            UID = uid;
        }

        public string UID { get; }

        public OpenExcelNumberingFormat NumberingFormat { get; set; }

        public OpenExcelFont Font { get; set; }

        public OpenExcelFill Fill { get; set; }

        public OpenExcelBorder Border { get; set; }
    }
}
