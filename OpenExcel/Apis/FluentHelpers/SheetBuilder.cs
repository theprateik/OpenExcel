using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcel.Abstractions.FluentApi;
using OpenExcel.Props;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel.Apis.FluentHelpers
{
    public class SheetBuilder
    {
        private readonly OpenExcelApi _api;
        private readonly IOpenExcelFluentApi _fluentApi;

        public  SheetBuilder(IOpenExcelFluentApi fluentApi)
        {
            _fluentApi = fluentApi;
            _api = _fluentApi.OpenExcelApi;
        }

        public SheetBuilder InsertSheetAs(string sheetName, OpenExcelSheetProperties sheetProperties = default)
        {
            _api.WriteEndSheet();

            _api.WriteStartSheet(sheetName, sheetProperties);

            return this;
        }

        public RowDataBuilder InsertRowData(List<string> cellValues, OpenExcelRowProperties rowProperties = default, EnumValue<CellValues> cellValueType = null)
        {
            _api.WriteRow(cellValues, rowProperties, cellValueType);

            return _fluentApi.RowDataBuilder;
        }

        public OpenExcelApi GetOpenExcelApi()
        {
            return _api;
        }
    }
}
