using OpenExcel.Abstractions.FluentApi;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel.Apis.FluentHelpers
{
    public class CellBuilder
    {
        private readonly OpenExcelApi _api;
        private readonly IOpenExcelFluentApi _fluentApi;

        public CellBuilder(IOpenExcelFluentApi fluentApi)
        {
            _fluentApi = fluentApi;
            _api = _fluentApi.OpenExcelApi;
        }

        public RowDataBuilder EndRow()
        {
            _api.WriteEndRow();

            return _fluentApi.RowDataBuilder;
        }

        public OpenExcelApi GetOpenExcelApi()
        {
            return _api;
        }
    }
}
