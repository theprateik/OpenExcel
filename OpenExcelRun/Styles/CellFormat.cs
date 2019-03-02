using OpenExcel.Styles;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcelRun.Styles
{
    public static class CellFormat
    {
        public static OpenExcelCellFormat C1 => new OpenExcelCellFormat("{8E11D0D2-E7CB-4512-88D3-767CACE014E5}") { Font = Font.F1 };

        public static OpenExcelCellFormat C2 => new OpenExcelCellFormat("{B409BC3E-036C-4BF1-8F47-C3EBA3A7DF13}") { Font = Font.F2, Fill = Fill.F1 };

        public static OpenExcelCellFormat C3 => new OpenExcelCellFormat("{AC8CFBB3-10EA-44CB-8E02-93D022F1D632}") { NumberingFormat = NumberFormat.N1 };

        public static OpenExcelCellFormat C4 => new OpenExcelCellFormat("{862B69A0-C87A-4CED-995E-E030EA1E2358}") { Font = Font.F2, Fill = Fill.F2 };

    }
}
