using OpenExcel.Props;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel.Abstractions.FluentApi
{
    public interface IRowBuilder
    {
        ICellBuilder InsertCell(string value, OpenExcelCellProperties cellProperties);
    }
}
