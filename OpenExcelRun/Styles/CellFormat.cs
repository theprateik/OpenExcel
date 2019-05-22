using OpenExcel.Styles;
using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenExcelRun.Styles
{
    public static class CellFormat
    {
        public static OpenExcelCellFormat C1 = new OpenExcelCellFormat(Guid.NewGuid().ToString()) { Font = Font.F1 };

        public static OpenExcelCellFormat C2 = new OpenExcelCellFormat(Guid.NewGuid().ToString()) { Font = Font.F2, Fill = Fill.F1 };

        public static OpenExcelCellFormat C3 = new OpenExcelCellFormat(Guid.NewGuid().ToString()) { NumberingFormat = NumberFormat.N1 };

        public static OpenExcelCellFormat C4 = new OpenExcelCellFormat(Guid.NewGuid().ToString()) { Font = Font.F2, Fill = Fill.F2 };

        public static OpenExcelCellFormat C5 = new OpenExcelCellFormat(Guid.NewGuid().ToString()) { Font = Font.F2, Fill = Fill.F2, NumberingFormat = NumberFormat.N1 };

        public static OpenExcelCellFormat C6 = new OpenExcelCellFormat(Guid.NewGuid().ToString()) { NumberingFormat = NumberFormat.N2 };

        public static OpenExcelCellFormat C7 = new OpenExcelCellFormat(Guid.NewGuid().ToString()) { NumberingFormat = NumberFormat.N3 };

        public static OpenExcelCellFormat C8 = new OpenExcelCellFormat(Guid.NewGuid().ToString()) { NumberingFormat = NumberFormat.N4 };

        public static OpenExcelCellFormat C9 = new OpenExcelCellFormat(Guid.NewGuid().ToString()) { Font = Font.F3, Fill = Fill.F1, Alignment = new OpenExcelAlignmentProperties() { Horizontal = HorizontalAlignmentValues.Center }};
    }
}
