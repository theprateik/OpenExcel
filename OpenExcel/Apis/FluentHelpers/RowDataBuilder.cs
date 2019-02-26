using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcel.Abstractions.FluentApi;
using OpenExcel.Models;
using OpenExcel.Props;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel.Apis.FluentHelpers
{
    public class RowDataBuilder
    {
        private readonly OpenExcelApi _api;
        private readonly IOpenExcelFluentApi _fluentApi;

        public RowDataBuilder(IOpenExcelFluentApi fluentApi)
        {
            _fluentApi = fluentApi;
            _api = _fluentApi.OpenExcelApi;
        }

        public SheetBuilder InsertSheetAs(string sheetName, OpenExcelSheetProperties sheetProperties = default)
        {
            _api.WriteEndSheet();

            _api.WriteStartSheet(sheetName, sheetProperties);

            return _fluentApi.SheetBuilder;
        }

        public RowDataBuilder InsertRowData(List<string> cellValues, OpenExcelRowProperties rowProperties = default, EnumValue<CellValues> cellValueType = null)
        {
            _api.WriteRow(cellValues, rowProperties, cellValueType);

            return this;
        }

        public RowDataBuilder InsertRowData<T>(T record, List<OpenExcelColumn<T>> columns, OpenExcelRowProperties rowProperties = default)
        {
            _api.WriteRow(record, columns, rowProperties);

            return this;
        }

        public RowDataBuilder InsertRowDataSet<T>(List<T> records, List<OpenExcelColumn<T>> columns, OpenExcelRowProperties rowProperties = default)
        {
            _api.WriteRowSet(records, columns, rowProperties);

            return this;
        }

        public RowBuilder CreateRow(OpenExcelRowProperties rowProperties)
        {
            _api.WriteStartRow(rowProperties);

            return _fluentApi.RowBuilder;
        }

        public void Complete()
        {
            _api.WriteEndSheet();
            _api.Close();
        }

        public OpenExcelApi GetOpenExcelApi()
        {
            return _api;
        }
    }
}
