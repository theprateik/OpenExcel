using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenExcel.Props
{
    public class OpenExcelSheetViewPaneProperties
    {
        public int XSplit { get; set; }
        public int YSplit { get; set; }
        public string TopLeftCell { get; set; }
        public EnumValue<PaneStateValues> State { get; set; }
    }
}
