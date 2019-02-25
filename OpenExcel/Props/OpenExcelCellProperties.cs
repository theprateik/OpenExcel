using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel.Props
{
    public class OpenExcelCellProperties
    {
        public EnumValue<CellValues> DataType { get; set; } = CellValues.String;
        public uint StyleIdx { get; set; } = 0;
    }
}
