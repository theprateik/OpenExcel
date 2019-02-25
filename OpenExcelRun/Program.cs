﻿using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcel;
using OpenExcel.Models;
using OpenExcel.Props;
using OpenExcel.Writers;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace OpenExcelRun
{
    class Program
    {
        static void Main()
        {
            var listPersons = new List<Person>();
           
            for (int i = 0; i <= 10000; i++)
            {
                listPersons.Add(new Person
                {
                    Age = 45,
                    DateOfBirth = DateTime.Now.AddYears(-25),
                    Name = "Sam Smith",
                    Income = 55600.28,
                    Children = new List<Child>
                    {
                        new Child { Name = "Tania", Age = 5, IsAdopted = false, IsHomeSchooled= true },
                        new Child { Name = "Tim", Age = 15, IsAdopted = false, IsHomeSchooled= false},
                        new Child { Name = "Sheena", Age = 26, IsAdopted = true, IsHomeSchooled= false},
                    }
                });
            }

            var columns = new List<OpenExcelColumn<Person>>
            {
                new OpenExcelColumn<Person>("Name", CellValues.SharedString, (x) => x.Name) { CellFormat = Styles.CellFormat.C1 },
                //new OpenExcelColumn<Person>("Name", CellValues.String, (x) => x.Name) ,
                new OpenExcelColumn<Person>("Age", CellValues.Number, (x) => x.Age.ToString()){ CellFormat = Styles.CellFormat.C3},
                new OpenExcelColumn<Person>("Date Of Birth", CellValues.SharedString, (x) => x.DateOfBirth.ToString()){ CellFormat = Styles.CellFormat.C1},
                new OpenExcelColumn<Person>("Income", CellValues.Number, (x) => x.Income.ToString()){ CellFormat = Styles.CellFormat.C1}
            };

            var childColumns = new List<OpenExcelColumn<Child>>
            {
                new OpenExcelColumn<Child>(string.Empty, CellValues.SharedString, (x) => string.Empty){ CellFormat = Styles.CellFormat.C2},
                new OpenExcelColumn<Child>("Name", CellValues.SharedString, (x) => x.Name){ CellFormat = Styles.CellFormat.C2} ,
                new OpenExcelColumn<Child>("Age", CellValues.Number, (x) => x.Age.ToString()){ CellFormat = Styles.CellFormat.C2},
                new OpenExcelColumn<Child>("Adopted?", CellValues.SharedString, (x) => x.IsAdopted ? "Yes" : "No")
                {
                    /*CellFormat = Styles.CellFormat.C2*/
                    CellFormatRule = (record, rowNum, colNum) => record.IsAdopted ? Styles.CellFormat.C4 : Styles.CellFormat.C2
                },
                new OpenExcelColumn<Child>("Home Schooled", CellValues.SharedString, (x) => x.IsHomeSchooled ? "Yes" : "No"){ CellFormat = Styles.CellFormat.C2}
            };

            //new ExcelExporter().CreateSpreadsheetWorkbook("D:\\Projects\\Temp\\Persons.xlsx", listPersons, columns);


            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            using (var openExcel = OpenExcelFactory.CreateOpenExcel("D:\\Projects\\Temp\\Persons2.xlsx"))
            {
                var sheetProperties = new OpenExcelSheetProperties
                {
                    OutlineProperties = new OpenExcelOutlineProperties { SummaryBelow = false }
                };

                openExcel.WriteStartSheet("Prateik Sheet", sheetProperties);
                openExcel.InsertHeader(columns);

                var childRowProperties = new OpenExcelRowProperties { OutlineLevel = 1 };

                foreach (var person in listPersons)
                {
                    openExcel.WriteRowSet(new List<Person> { person }, columns);
                    openExcel.InsertHeader(childColumns, childRowProperties);
                    openExcel.WriteRowSet(person.Children, childColumns, childRowProperties);
                    openExcel.WriteRow(new List<string> { string.Empty }, childRowProperties, CellValues.SharedString);
                }
                openExcel.WriteEndSheet();

                openExcel.WriteStartSheet("Ronas Sheet");
                openExcel.WriteRowSet(listPersons, columns);
                openExcel.WriteEndSheet();

                openExcel.Close();
            }

            stopwatch.Stop();
            Console.WriteLine("Time elapsed: {0}", stopwatch.Elapsed);
        }
    }
}
