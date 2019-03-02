using OpenExcel.Apis;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel.Abstractions.FluentApi
{
    public interface IOpenExcelBuilder : IDisposable
    {
        IExcelBuilder CreateExcelAs(string filePath);
    }
}
