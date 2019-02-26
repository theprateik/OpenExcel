using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcel.Abstractions.FluentApi;
using OpenExcel.Apis;
using OpenExcel.Models;
using OpenExcel.Props;
using OpenExcel.Writers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OpenExcel
{
    public static class OpenExcelExtensions
    {
        public static void InsertHeader<T>(this OpenExcelApi writer, List<OpenExcelColumn<T>> columns, OpenExcelRowProperties rowProperties = default)
        {
            writer.WriteRow(columns.Select(x => x.Name).ToList(), rowProperties, CellValues.SharedString);
        }
    }
}
