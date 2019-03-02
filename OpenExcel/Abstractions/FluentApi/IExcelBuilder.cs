using OpenExcel.Props;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel.Abstractions.FluentApi
{
    public interface IExcelBuilder
    {
        ISheetBuilder InsertSheetAs(string sheetName = default, OpenExcelSheetProperties sheetProperties = default);
        void Complete();
    }
}
