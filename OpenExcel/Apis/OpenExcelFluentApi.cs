using OpenExcel.Abstractions.FluentApi;
using OpenExcel.Apis.FluentHelpers;
using OpenExcel.Props;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel.Apis
{
    public class OpenExcelFluentApi : IOpenExcelFluentApi, IOpenExcelBuilder, ISheetBuilder
    {
        private OpenExcelFluentApi()
        {

        }

        public SheetBuilder SheetBuilder {get; private set;}
        public RowDataBuilder RowDataBuilder {get; private set;}
        public RowBuilder RowBuilder {get; private set;}
        public CellBuilder CellBuilder {get; private set; }
        public OpenExcelApi OpenExcelApi { get; private set; }

        public ISheetBuilder CreateExcelAs(string filePath)
        {
            OpenExcelApi = new OpenExcelApi(filePath);
            SheetBuilder = new SheetBuilder(this);
            RowDataBuilder = new RowDataBuilder(this);
            RowBuilder = new RowBuilder(this);
            CellBuilder = new CellBuilder(this);

            return this;
        }

        public SheetBuilder InsertSheetAs(string sheetName, OpenExcelSheetProperties sheetProperties = default)
        {
            OpenExcelApi.WriteStartSheet(sheetName, sheetProperties);

            var sheetbuilder = new SheetBuilder(this);

            return sheetbuilder;
        }

        public OpenExcelApi GetOpenExcelApi()
        {
            return OpenExcelApi;
        }

        public static IOpenExcelBuilder CreateOpenExcelBuilder()
        {
            return new OpenExcelFluentApi();
        }
    }
}
