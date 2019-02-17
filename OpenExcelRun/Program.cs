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
            var listPersons = new List<Person>();

            for (int i = 0; i <= 1000000; i++)
            {
                listPersons.Add(new Person
                {
                    Age = 45,
                    DateOfBirth = DateTime.Now.AddYears(-25),
                    Name = "Sam Smith",
                    Income = 55600.28
                });
            }

            var columns = new List<OpenExcelColumn<Person>>
            {
                new OpenExcelColumn<Person>("Name", CellValues.String, (x) => x.Name),
                new OpenExcelColumn<Person>("Age", CellValues.Number, (x) => x.Age.ToString()),
                new OpenExcelColumn<Person>("Date Of Birth", CellValues.String, (x) => x.DateOfBirth.ToString()),
                new OpenExcelColumn<Person>("Income", CellValues.Number, (x) => x.Income.ToString())
            };

            new ExcelExporter().CreateSpreadsheetWorkbook("D:\\Projects\\Temp\\Persons.xlsx", listPersons, columns);
        }
    }

}
