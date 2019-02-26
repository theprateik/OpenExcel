using OpenExcel.Apis;
using OpenExcel.Apis.FluentHelpers;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel.Abstractions.FluentApi
{
    public interface IOpenExcelFluentApi
    {
        SheetBuilder SheetBuilder { get;  }
        RowDataBuilder RowDataBuilder { get;  }
        RowBuilder RowBuilder { get;  }
        CellBuilder CellBuilder { get; }
        OpenExcelApi OpenExcelApi { get; }
    }
}
