using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel
{
    public class OpenExcelColumn<T>
    {
        public OpenExcelColumn(string name, CellValues cellValueType, Func<T, string> selector)
        {
            Name = name;
            Selector = selector;
            CellValueType = cellValueType;
        }

        public string Name { get; set; }

        public EnumValue<CellValues> CellValueType { get; set; }

        public Func<T, string> Selector { get; set; }
    }
}
