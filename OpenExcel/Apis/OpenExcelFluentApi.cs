using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcel.Abstractions.FluentApi;
using OpenExcel.Models;
using OpenExcel.Props;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenExcel.Styles;

namespace OpenExcel.Apis
{
    public class OpenExcelFluentApi : IOpenExcelBuilder, IExcelBuilder, ISheetBuilder, IRowBuilder, ICellBuilder
    {
        private OpenExcelApi _api;

        private OpenExcelFluentApi()
        {

        }

        public IExcelBuilder CreateExcelAs(string filePath)
        {
            _api = new OpenExcelApi(filePath);

            return this;
        }

        ISheetBuilder IExcelBuilder.InsertSheetAs(string sheetName, OpenExcelSheetProperties sheetProperties, OpenExcelSheetViewProperties sheetViewProperties, OpenExcelSheetFormatProperties sheetFormatProperties)
        {
            _api.WriteStartSheet(sheetName, sheetProperties, sheetViewProperties, sheetFormatProperties);

            return this;
        }

        ISheetBuilder IExcelBuilder.InsertSheetWithFirstRowFrozenAs(string sheetName, OpenExcelSheetProperties sheetProperties, OpenExcelSheetFormatProperties sheetFormatProperties)
        {
            var sheetViewProperties = new OpenExcelSheetViewProperties
            {
                PaneProperties = new OpenExcelSheetViewPaneProperties
                {
                    XSplit = 0, YSplit = 1, TopLeftCell = "A2", State = PaneStateValues.FrozenSplit
                }
            };

            (this as IExcelBuilder).InsertSheetAs(sheetName, sheetProperties, sheetViewProperties, sheetFormatProperties);

            return this;
        }

        ISheetBuilder ISheetBuilder.InsertSheetAs(string sheetName, OpenExcelSheetProperties sheetProperties, OpenExcelSheetViewProperties sheetViewProperties, OpenExcelSheetFormatProperties sheetFormatProperties)
        {
            _api.WriteEndSheet();

            (this as IExcelBuilder).InsertSheetAs(sheetName, sheetProperties, sheetViewProperties, sheetFormatProperties);

            return this;
        }

        ISheetBuilder ISheetBuilder.InsertSheetWithFirstRowFrozenAs(string sheetName, OpenExcelSheetProperties sheetProperties, OpenExcelSheetFormatProperties sheetFormatProperties)
        {
            _api.WriteEndSheet();

            (this as IExcelBuilder).InsertSheetWithFirstRowFrozenAs(sheetName, sheetProperties, sheetFormatProperties);

            return this;
        }

        public ISheetBuilder InsertRowData(List<string> cellValues, OpenExcelRowProperties rowProperties = null, EnumValue<CellValues> cellValueType = null)
        {
            _api.WriteRow(cellValues, rowProperties, cellValueType);

            return this;
        }

        public ISheetBuilder InsertRowData<T>(T record, List<OpenExcelColumn<T>> columns, OpenExcelRowProperties rowProperties = null)
        {
            _api.WriteRow(record, columns, rowProperties);

            return this;
        }

        public ISheetBuilder InsertRowDataSet<T>(IEnumerable<T> records, List<OpenExcelColumn<T>> columns, OpenExcelRowProperties rowProperties = null)
        {
            _api.WriteRowSet(records, columns, rowProperties);

            return this;
        }

        public ISheetBuilder InsertHeaderRow<T>(List<OpenExcelColumn<T>> columns, OpenExcelCellFormat cellFormat, OpenExcelRowProperties rowProperties = null)
        {
            var headerColumns = columns.Select(x => new OpenExcelColumn<string>(x.Name, CellValues.SharedString, y => x.Name) { CellFormat = x.HeaderCellFormat ?? cellFormat }).ToList();

            (this as ISheetBuilder).InsertRowData("", headerColumns, rowProperties);

            return this;
        }

        public ISheetBuilder InsertEmptyRow()
        {
            (this as ISheetBuilder).InsertRowData(default);

            return this;
        }

        public IRowBuilder CreateRow(OpenExcelRowProperties rowProperties)
        {
            _api.WriteStartRow(rowProperties);

            return this;
        }

        public ICellBuilder InsertCell(string value, OpenExcelCellProperties cellProperties)
        {
            _api.WriteCell(value, cellProperties);

            return this;
        }

        public ISheetBuilder EndRow()
        {
            _api.WriteEndRow();

            return this;
        }

        void IExcelBuilder.Complete()
        {
            (this as IExcelBuilder).InsertSheetAs("Sheet1");
            _api.WriteEndSheet();
            _api.Close();
        }

        void ISheetBuilder.Complete()
        {
            _api.WriteEndSheet();
            _api.Close();
        }

        public void Dispose()
        {
            _api.Dispose();
        }

        public static IOpenExcelBuilder CreateOpenExcelBuilder()
        {
            return new OpenExcelFluentApi();
        }
    }
}
