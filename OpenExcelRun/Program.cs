using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcel;
using OpenExcel.Abstractions.FluentApi;
using OpenExcel.Apis;
using OpenExcel.Models;
using OpenExcel.Props;
using OpenExcelRun.CreateExcel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace OpenExcelRun
{
    class Program
    {
        static void Main()
        {
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            UsingFluentApi.Run();


            stopwatch.Stop();
            Console.WriteLine("Time elapsed: {0}", stopwatch.Elapsed);

            stopwatch.Reset();
            stopwatch.Start();

            UsingFluentApi.Run();

            stopwatch.Stop();
            Console.WriteLine("Time elapsed: {0}", stopwatch.Elapsed);

            stopwatch.Reset();
            stopwatch.Start();


            ByOpenExcelApi.Run();


            stopwatch.Stop();
            Console.WriteLine("Time elapsed: {0}", stopwatch.Elapsed);

            //var fluent = OpenExcelFluentApi.CreateOpenExcelBuilder();

            //fluent.CreateExcelAs("D:\\Projects\\Temp\\Persons3.xlsx")
            //    .InsertSheetAs("Prateik")
            //    .InsertRowData(new List<string> {"Ram", "Shyam" }, cellValueType: CellValues.String)
            //    .InsertSheetAs("Prateik 123")
            //    .InsertRowData(new List<string> { "Ronas", "Dinesh" }, cellValueType: CellValues.String)
            //    .Complete();


            //fluent.CreateExcelAs("D:\\Projects\\Temp\\Persons4.xlsx")
            //    .InsertSheetAs("Prateik")
            //    .CreateRow(new OpenExcelRowProperties())
            //    .InsertCell("test", new OpenExcelCellProperties { DataType = CellValues.String })
            //    .InsertCell("test", new OpenExcelCellProperties { DataType = CellValues.String })
            //    .InsertCell("test", new OpenExcelCellProperties { DataType = CellValues.String })
            //    .InsertCell("test", new OpenExcelCellProperties { DataType = CellValues.String })
            //    .InsertCell("test", new OpenExcelCellProperties { DataType = CellValues.String })
            //    .InsertCell("test", new OpenExcelCellProperties { DataType = CellValues.String })
            //    .InsertCell("test", new OpenExcelCellProperties { DataType = CellValues.String })
            //    .EndRow()
            //    .Complete();

            //fluent.CreateExcelAs("D:\\Projects\\Temp\\Persons5.xlsx")
            //    .InsertSheetAs("")
            //    .CreateRow(new OpenExcelRowProperties())
            //    .InsertCell("test", new OpenExcelCellProperties { DataType = CellValues.String })
            //    .InsertCell("test", new OpenExcelCellProperties { DataType = CellValues.String })
            //    .InsertCell("test", new OpenExcelCellProperties { DataType = CellValues.String })
            //    .InsertCell("test", new OpenExcelCellProperties { DataType = CellValues.String })
            //    .InsertCell("test", new OpenExcelCellProperties { DataType = CellValues.String })
            //    .InsertCell("test", new OpenExcelCellProperties { DataType = CellValues.String })
            //    .InsertCell("test", new OpenExcelCellProperties { DataType = CellValues.String })
            //    .EndRow()
            //    .InsertEmptyRow()
            //    .CreateRow(new OpenExcelRowProperties())
            //    .InsertCell("test", new OpenExcelCellProperties { DataType = CellValues.String })
            //    .InsertCell("test", new OpenExcelCellProperties { DataType = CellValues.String })
            //    .InsertCell("test", new OpenExcelCellProperties { DataType = CellValues.String })
            //    .InsertCell("test", new OpenExcelCellProperties { DataType = CellValues.String })
            //    .InsertCell("test", new OpenExcelCellProperties { DataType = CellValues.String })
            //    .InsertCell("test", new OpenExcelCellProperties { DataType = CellValues.String })
            //    .InsertCell("test", new OpenExcelCellProperties { DataType = CellValues.String })
            //    .EndRow()
            //    //.InsertSheetAs()
            //    //.CreateRow(new OpenExcelRowProperties())
            //    //.InsertSheetAs("test")
            //    .Complete();




        }
    }
}
