using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenExcel.Styles
{
    public class OpenExcelAlignmentProperties
    {
        public HorizontalAlignmentValues Horizontal { get; set; }
        public VerticalAlignmentValues Vertical { get; set; }
        public bool WrapText { get; set; }
        public uint TextRotation { get; set; }
    }
}
