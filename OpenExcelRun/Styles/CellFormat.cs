using OpenExcel.Styles;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcelRun.Styles
{
    public static class CellFormat
    {
        public static OpenExcelCellFormat C1 = new OpenExcelCellFormat("5C973B80-5800-4A03-80B7-C0569E85DDD3") { Font = Font.F1 };

        public static OpenExcelCellFormat C2 = new OpenExcelCellFormat("16EA7C41-C7DF-4B9C-A552-F5AC75DF492F") { Font = Font.F2, Fill = Fill.F1 };

        public static OpenExcelCellFormat C3 = new OpenExcelCellFormat(Guid.NewGuid().ToString()) { NumberingFormat = NumberFormat.N1 };

        public static OpenExcelCellFormat C4 = new OpenExcelCellFormat(Guid.NewGuid().ToString()) { Font = Font.F2, Fill = Fill.F2 };

    }
}
