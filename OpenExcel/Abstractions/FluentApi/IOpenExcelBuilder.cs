using OpenExcel.Apis;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel.Abstractions.FluentApi
{
    public interface IOpenExcelBuilder
    {
        ISheetBuilder CreateExcelAs(string filePath);
    }
}
