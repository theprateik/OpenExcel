using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcel.Models;
using OpenExcel.Props;
using System;
using System.Collections.Generic;
using System.Text;
using OpenExcel.Styles;

namespace OpenExcel.Abstractions.FluentApi
{
    public interface ISheetBuilder
    {
        ISheetBuilder InsertSheetAs(string sheetName = default, OpenExcelSheetProperties sheetProperties = default, OpenExcelSheetViewProperties sheetViewProperties = default, OpenExcelSheetFormatProperties sheetFormatProperties = default);
        ISheetBuilder InsertSheetWithFirstRowFrozenAs(string sheetName = default, OpenExcelSheetProperties sheetProperties = default, OpenExcelSheetFormatProperties sheetFormatProperties = default);
        ISheetBuilder InsertRowData(List<string> cellValues, OpenExcelRowProperties rowProperties = default, EnumValue<CellValues> cellValueType = default);
        ISheetBuilder InsertRowData<T>(T record, List<OpenExcelColumn<T>> columns, OpenExcelRowProperties rowProperties = default);
        ISheetBuilder InsertRowDataSet<T>(IEnumerable<T> records, List<OpenExcelColumn<T>> columns, OpenExcelRowProperties rowProperties = default);
        ISheetBuilder InsertEmptyRow();
        ISheetBuilder InsertHeaderRow<T>(List<OpenExcelColumn<T>> columns, OpenExcelCellFormat cellFormat = default, OpenExcelRowProperties rowProperties = default);
        IRowBuilder CreateRow(OpenExcelRowProperties rowProperties);
        void Complete();
    }
}
