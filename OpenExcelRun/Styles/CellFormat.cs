using OpenExcel.Styles;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcelRun.Styles
{
    public static class CellFormat
    {
        public static OpenExcelCellFormat C1 = new OpenExcelCellFormat(Guid.NewGuid().ToString()) { Font = Font.F1 };

        public static OpenExcelCellFormat C2 = new OpenExcelCellFormat(Guid.NewGuid().ToString()) { Font = Font.F2, Fill = Fill.F1 };

        public static OpenExcelCellFormat C3 = new OpenExcelCellFormat(Guid.NewGuid().ToString()) { NumberingFormat = NumberFormat.N1 };

        public static OpenExcelCellFormat C4 = new OpenExcelCellFormat(Guid.NewGuid().ToString()) { Font = Font.F2, Fill = Fill.F2 };

    }
}
