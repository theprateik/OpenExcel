using OpenExcel.Props;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel.Abstractions.FluentApi
{
    public interface IExcelBuilder
    {
        ISheetBuilder InsertSheetAs(string sheetName = default, OpenExcelSheetProperties sheetProperties = default, OpenExcelSheetViewProperties sheetViewProperties = default, OpenExcelSheetFormatProperties sheetFormatProperties = default);
        ISheetBuilder InsertSheetWithFirstRowFrozenAs(string sheetName = default, OpenExcelSheetProperties sheetProperties = default, OpenExcelSheetFormatProperties sheetFormatProperties = default);
        void Complete();
    }
}
