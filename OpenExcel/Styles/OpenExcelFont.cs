using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel.Styles
{
    public class OpenExcelFont
    {
        public OpenExcelFont(string uid)
        {
            UID = uid;
        }
        public string UID { get; }
        public uint? FontFamilyNumbering { get; set; }
        public string FontName { get; set; }
        public HexBinaryValue Color { get; set; }
        public uint? FontSize { get; set; }
        public bool Italic { get; set; }
        public bool Bold { get; set; }        
    }
}
