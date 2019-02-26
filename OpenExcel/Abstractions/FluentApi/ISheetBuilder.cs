using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcel.Apis.FluentHelpers;
using OpenExcel.Models;
using OpenExcel.Props;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel.Abstractions.FluentApi
{
    public interface ISheetBuilder
    {
        SheetBuilder InsertSheetAs(string sheetName, OpenExcelSheetProperties sheetProperties = default);
    }
}
