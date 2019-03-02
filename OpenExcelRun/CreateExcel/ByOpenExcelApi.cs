using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcel;
using OpenExcel.Models;
using OpenExcel.Props;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcelRun.CreateExcel
{
    public static class ByOpenExcelApi
    {
        public static void Run()
        {
            var listPersons = new List<Person>();

            for (int i = 0; i <= 100000; i++)
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



            using (var api = OpenExcelFactory.CreateOpenExcelApi("D:\\Projects\\Temp\\Persons2.xlsx"))
            {
                var sheetProperties = new OpenExcelSheetProperties
                {
                    OutlineProperties = new OpenExcelOutlineProperties { SummaryBelow = false }
                };

                api.WriteStartSheet("Prateik Sheet", sheetProperties);
                api.InsertHeader(columns);

                var childRowProperties = new OpenExcelRowProperties { OutlineLevel = 1 };

                foreach (var person in listPersons)
                {
                    api.WriteRowSet(new List<Person> { person }, columns);
                    api.InsertHeader(childColumns, childRowProperties);
                    api.WriteRowSet(person.Children, childColumns, childRowProperties);
                    api.WriteRow(new List<string> { string.Empty }, childRowProperties, CellValues.SharedString);
                }
                api.WriteEndSheet();

                api.WriteStartSheet("Ronas Sheet");
                api.WriteRowSet(listPersons, columns);
                api.WriteEndSheet();

                api.Close();
            }

            
        }
    }
}
