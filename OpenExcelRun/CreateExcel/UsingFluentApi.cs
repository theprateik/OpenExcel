using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcel.Apis;
using OpenExcel.Models;
using OpenExcel.Props;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OpenExcelRun.CreateExcel
{
    public static class UsingFluentApi
    {
        public static void Run()
        {
            var columns = new List<OpenExcelColumn<Person>>
            {
                new OpenExcelColumn<Person>("Name", CellValues.SharedString, (x) => x.Name) { CellFormat = Styles.CellFormat.C1, HeaderCellFormat = Styles.CellFormat.C9 },
                //new OpenExcelColumn<Person>("Name", CellValues.String, (x) => x.Name) ,
                new OpenExcelColumn<Person>("Age", CellValues.Date, (x) => x.Age.ToString()){ CellFormat = Styles.CellFormat.C3},
                new OpenExcelColumn<Person>("Date Of Birth", CellValues.Date, (x) => "43101"){ CellFormat = Styles.CellFormat.C8},
                new OpenExcelColumn<Person>("Income", CellValues.Number, (x) => x.Income.ToString()){ CellFormat = Styles.CellFormat.C7},
                new OpenExcelColumn<Person>("Income", CellValues.Number, (x) => x.Income.ToString()){ CellFormat = Styles.CellFormat.C7},
                new OpenExcelColumn<Person>("Income", CellValues.Number, (x) => x.Income.ToString()){ CellFormat = Styles.CellFormat.C7},
                new OpenExcelColumn<Person>("Income", CellValues.Number, (x) => x.Income.ToString()){ CellFormat = Styles.CellFormat.C7},
                new OpenExcelColumn<Person>("Income", CellValues.Number, (x) => x.Income.ToString()){ CellFormat = Styles.CellFormat.C7},
                new OpenExcelColumn<Person>("Income", CellValues.Number, (x) => x.Income.ToString()){ CellFormat = Styles.CellFormat.C7},
                new OpenExcelColumn<Person>("Income", CellValues.Number, (x) => x.Income.ToString()){ CellFormat = Styles.CellFormat.C7},
                new OpenExcelColumn<Person>("Income", CellValues.Number, (x) => x.Income.ToString()){ CellFormat = Styles.CellFormat.C7},
                new OpenExcelColumn<Person>("Income", CellValues.Number, (x) => x.Income.ToString()){ CellFormat = Styles.CellFormat.C7},
                new OpenExcelColumn<Person>("Income", CellValues.Number, (x) => x.Income.ToString()){ CellFormat = Styles.CellFormat.C7},
                new OpenExcelColumn<Person>("Income", CellValues.Number, (x) => x.Income.ToString()){ CellFormat = Styles.CellFormat.C7},
                new OpenExcelColumn<Person>("Income", CellValues.Number, (x) => x.Income.ToString()){ CellFormat = Styles.CellFormat.C7},
                new OpenExcelColumn<Person>("Income", CellValues.Number, (x) => x.Income.ToString()){ CellFormat = Styles.CellFormat.C7},
                new OpenExcelColumn<Person>("Income", CellValues.Number, (x) => x.Income.ToString()){ CellFormat = Styles.CellFormat.C7},
                new OpenExcelColumn<Person>("Income", CellValues.Number, (x) => x.Income.ToString()){ CellFormat = Styles.CellFormat.C7},
                new OpenExcelColumn<Person>("Income", CellValues.Number, (x) => x.Income.ToString()){ CellFormat = Styles.CellFormat.C7},
                new OpenExcelColumn<Person>("Income", CellValues.Number, (x) => x.Income.ToString()){ CellFormat = Styles.CellFormat.C7},
                new OpenExcelColumn<Person>("Income", CellValues.Number, (x) => x.Income.ToString()){ CellFormat = Styles.CellFormat.C7}
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
                new OpenExcelColumn<Child>("Home Schooled", CellValues.SharedString, (x) => x.IsHomeSchooled ? "Yes" : "No"){ CellFormat = Styles.CellFormat.C2},
                new OpenExcelColumn<Child>("Name", CellValues.SharedString, (x) => x.Name){ CellFormat = Styles.CellFormat.C2} ,
                new OpenExcelColumn<Child>("Name", CellValues.SharedString, (x) => x.Name){ CellFormat = Styles.CellFormat.C2} ,
                new OpenExcelColumn<Child>("Name", CellValues.SharedString, (x) => x.Name){ CellFormat = Styles.CellFormat.C2} ,
                new OpenExcelColumn<Child>("Name", CellValues.SharedString, (x) => x.Name){ CellFormat = Styles.CellFormat.C2} ,
                new OpenExcelColumn<Child>("Name", CellValues.SharedString, (x) => x.Name){ CellFormat = Styles.CellFormat.C2} ,
                new OpenExcelColumn<Child>("Name", CellValues.SharedString, (x) => x.Name){ CellFormat = Styles.CellFormat.C2} 
            };


            using (var fluent = OpenExcelFluentApi.CreateOpenExcelBuilder())
            {
                var sheetBuilder = fluent.CreateExcelAs("E:\\Projects\\Temp\\Persons67.xlsx")
                    .InsertSheetWithFirstRowFrozenAs("Prateik",
                        new OpenExcelSheetProperties
                            {OutlineProperties = new OpenExcelOutlineProperties {SummaryBelow = false}})
                    .InsertHeaderRow(columns, Styles.CellFormat.C9);

                var childRowProperties = new OpenExcelRowProperties { OutlineLevel = 1 };
                var listPersons = GenerateData();
                foreach (var person in listPersons)
                {
                    sheetBuilder = sheetBuilder
                        .InsertRowData(person, columns)
                        .InsertHeaderRow(childColumns, Styles.CellFormat.C1, childRowProperties)
                        .InsertRowDataSet(person.Children, childColumns, childRowProperties)
                        //.InsertEmptyRow();
                        .InsertRowData(new List<string> { string.Empty }, childRowProperties, CellValues.SharedString);
                }

                sheetBuilder.Complete();
                //sheetBuilder.InsertSheetAs("Ronas Sheet")
                //    .InsertRowDataSet(listPersons, columns)
                //    .Complete();
            }
        }

        public static IEnumerable<Person> GenerateData()
        {
            for (int i = 0; i <= 100000; i++)
            {
                yield return new Person
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
                };
            }
        }
    }
}
