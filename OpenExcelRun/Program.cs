using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcel;
using System;
using System.Collections.Generic;

namespace OpenExcelRun
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            new ExcelExporter().CreateSpreadsheetWorkbook("D:\\Projects\\Temp\\test.xlsx");
        }


    }

}
