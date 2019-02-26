using OpenExcel.Abstractions.FluentApi;
using OpenExcel.Props;

namespace OpenExcel.Apis.FluentHelpers
{
    public class RowBuilder
    {
        private readonly OpenExcelApi _api;
        private readonly IOpenExcelFluentApi _fluentApi;

        public RowBuilder(IOpenExcelFluentApi fluentApi)
        {
            _fluentApi = fluentApi;
            _api = _fluentApi.OpenExcelApi;
        }

        public CellBuilder InsertCell(string cellValue, OpenExcelCellProperties cellProperties)
        {
            _api.WriteCell(cellValue, cellProperties);

            return _fluentApi.CellBuilder;
        }

        public OpenExcelApi GetOpenExcelApi()
        {
            return _api;
        }
    }
}
