using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel
{
    public class ExcelExporter
    {
        public void CreateSpreadsheetWorkbook1<T>(string filePath, List<T> data, List<OpenExcelColumn<T>> columns)
        {
            using (SpreadsheetDocument xl = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                OpenXmlWriter oxw;
                int rowCounter = 0;

                xl.AddWorkbookPart();
                AddStyleSheetOld(xl);
                WorksheetPart wsp = xl.WorkbookPart.AddNewPart<WorksheetPart>();


                oxw = OpenXmlWriter.Create(wsp);
                {

                    oxw.WriteStartElement(new Worksheet());
                    {
                        oxw.WriteStartElement(new SheetData());
                        {
                            AddData(oxw, data, columns);
                        }
                        // this is for SheetData
                        oxw.WriteEndElement();
                    }
                    oxw.WriteEndElement(); // this is for Worksheet
                }
                oxw.Close();

                oxw = OpenXmlWriter.Create(xl.WorkbookPart);
                {
                    oxw.WriteStartElement(new Workbook());
                    {
                        oxw.WriteStartElement(new Sheets());
                        {
                            // you can use object initialisers like this only when the properties
                            // are actual properties. SDK classes sometimes have property-like properties
                            // but are actually classes. For example, the Cell class has the CellValue
                            // "property" but is actually a child class internally.
                            // If the properties correspond to actual XML attributes, then you're fine.
                            oxw.WriteElement(new Sheet()
                            {
                                Name = "Sheet1",
                                SheetId = 1,
                                Id = xl.WorkbookPart.GetIdOfPart(wsp)
                            });

                            oxw.WriteElement(new Sheet()
                            {
                                Name = "Sheet2",
                                SheetId = 2,
                                Id = xl.WorkbookPart.GetIdOfPart(wsp)
                            });
                        }
                        // this is for Sheets
                        oxw.WriteEndElement();
                    }
                    // this is for Workbook
                    oxw.WriteEndElement();
                }
                oxw.Close();

                xl.Close();
            }
        }

        public void CreateSpreadsheetWorkbook<T>(string filePath, List<T> data, List<OpenExcelColumn<T>> columns)
        {
            using (SpreadsheetDocument xl = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                OpenXmlWriter oxw;
                int rowCounter = 0;

                xl.AddWorkbookPart();
                AddStyleSheetOld(xl);
                WorksheetPart wsp = xl.WorkbookPart.AddNewPart<WorksheetPart>();

                var wbWriter = OpenXmlWriter.Create(xl.WorkbookPart);
                wbWriter.WriteStartElement(new Workbook());
                wbWriter.WriteStartElement(new Sheets());
                {
                    // you can use object initialisers like this only when the properties
                    // are actual properties. SDK classes sometimes have property-like properties
                    // but are actually classes. For example, the Cell class has the CellValue
                    // "property" but is actually a child class internally.
                    // If the properties correspond to actual XML attributes, then you're fine.
                    wbWriter.WriteElement(new Sheet()
                    {
                        Name = "Sheet1",
                        SheetId = 1,
                        Id = xl.WorkbookPart.GetIdOfPart(wsp)
                    });

                    //wbWriter.WriteElement(new Sheet()
                    //{
                    //    Name = "Sheet2",
                    //    SheetId = 2,
                    //    Id = xl.WorkbookPart.GetIdOfPart(wsp)
                    //});
                }



                oxw = OpenXmlWriter.Create(wsp);
                {

                    oxw.WriteStartElement(new Worksheet());
                    {
                        oxw.WriteStartElement(new SheetData());
                        {
                            AddData(oxw, data, columns);
                        }
                        // this is for SheetData
                        oxw.WriteEndElement();
                    }
                    oxw.WriteEndElement(); // this is for Worksheet
                }

                
                // this is for Sheets
                wbWriter.WriteEndElement();

                oxw.Close();

                //oxw = OpenXmlWriter.Create(xl.WorkbookPart);
                //{
                //    oxw.WriteStartElement(new Workbook());
                //    {
                //        oxw.WriteStartElement(new Sheets());
                //        {
                //            // you can use object initialisers like this only when the properties
                //            // are actual properties. SDK classes sometimes have property-like properties
                //            // but are actually classes. For example, the Cell class has the CellValue
                //            // "property" but is actually a child class internally.
                //            // If the properties correspond to actual XML attributes, then you're fine.
                //            oxw.WriteElement(new Sheet()
                //            {
                //                Name = "Sheet1",
                //                SheetId = 1,
                //                Id = xl.WorkbookPart.GetIdOfPart(wsp)
                //            });

                //            oxw.WriteElement(new Sheet()
                //            {
                //                Name = "Sheet2",
                //                SheetId = 2,
                //                Id = xl.WorkbookPart.GetIdOfPart(wsp)
                //            });
                //        }
                //        // this is for Sheets
                //        oxw.WriteEndElement();
                //    }
                // this is for Workbook
                wbWriter.WriteEndElement();
                //}
                wbWriter.Close();

                xl.Close();
            }
        }


        public void AddStyleSheetOld(SpreadsheetDocument document)
        {
            var ss = document.WorkbookPart.AddNewPart<WorkbookStylesPart>();

            var ssWriter = OpenXmlWriter.Create(ss);
            {
                ssWriter.WriteStartElement(new Stylesheet());
                {
                    ssWriter.WriteStartElement(new NumberingFormats());
                    {
                        //ssWriter.WriteElement(new NumberingFormat());
                        ssWriter.WriteElement(new NumberingFormat() { NumberFormatId = 164U, FormatCode = "mm/dd/yyyy hh:mm:ss" });
                    }
                    ssWriter.WriteEndElement();

                    ssWriter.WriteStartElement(new Fonts());
                    {
                        ssWriter.WriteStartElement(new Font());
                        {
                        }
                        ssWriter.WriteEndElement();

                        ssWriter.WriteStartElement(new Font());
                        {
                            ssWriter.WriteElement(new Italic());
                            ssWriter.WriteElement(new FontSize() { Val = 11 });
                        }
                        ssWriter.WriteEndElement();

                        ssWriter.WriteStartElement(new Font());
                        {
                            ssWriter.WriteElement(new Bold());
                            ssWriter.WriteElement(new FontSize() { Val = 14 });
                        }
                        ssWriter.WriteEndElement();
                    }
                    ssWriter.WriteEndElement();

                    ssWriter.WriteStartElement(new Fills());
                    {
                        ssWriter.WriteStartElement(new Fill());
                        {
                            ssWriter.WriteElement(new PatternFill() { PatternType = PatternValues.None });
                        }
                        ssWriter.WriteEndElement();

                        ssWriter.WriteStartElement(new Fill());
                        {
                            ssWriter.WriteElement(new PatternFill() { PatternType = PatternValues.DarkGray });
                        }
                        ssWriter.WriteEndElement();
                    }
                    ssWriter.WriteEndElement();

                    ssWriter.WriteStartElement(new Borders());
                    {
                        ssWriter.WriteStartElement(new Border());
                        {
                            ssWriter.WriteElement(new LeftBorder());
                            ssWriter.WriteElement(new RightBorder());
                            ssWriter.WriteElement(new TopBorder());
                            ssWriter.WriteElement(new BottomBorder());
                            ssWriter.WriteElement(new DiagonalBorder());
                        }
                        ssWriter.WriteEndElement();
                    }
                    ssWriter.WriteEndElement();

                    ssWriter.WriteStartElement(new CellStyleFormats() { Count = 1 });
                    {
                        ssWriter.WriteElement(new CellFormat() { FontId = 0U, FillId = 0U, BorderId = 0U });
                    }
                    ssWriter.WriteEndElement();

                    ssWriter.WriteStartElement(new CellFormats() /*{ Count = 1 }*/);
                    {
                        ssWriter.WriteElement(new CellFormat() { FontId = 0U, FillId = 0U, BorderId = 0U });
                        ssWriter.WriteElement(new CellFormat() { FontId = 1U, FillId = 1U, BorderId = 0U });
                        ssWriter.WriteElement(new CellFormat() { FontId = 2U, FillId = 0U, BorderId = 0U });
                        ssWriter.WriteElement(new CellFormat() { NumberFormatId = 164U, FontId = 2U, FillId = 1U, BorderId = 0U, ApplyNumberFormat = true, FormatId = 0, });
                        //ssWriter.WriteElement(new CellFormat() { FontId = 1, FillId = 0 });
                    }
                    ssWriter.WriteEndElement();
                }
                ssWriter.WriteEndElement();
            }
            ssWriter.Close();
        }

        public void AddData<T>(OpenXmlWriter oxw, List<T> data, List<OpenExcelColumn<T>> columns)
        {
            List<OpenXmlAttribute> oxa;
            for (int i = 0; i < data.Count; i++)
            {
                //rowCounter++;
                oxa = new List<OpenXmlAttribute>();
                // this is the row index
                //oxa.Add(new OpenXmlAttribute("r", null, i.ToString()));
                if (i > 5 && i < 15)
                {
                    oxa.Add(new OpenXmlAttribute("outlineLevel", string.Empty, "1"));
                }
                oxw.WriteStartElement(new Row(), oxa);

                foreach (var column in columns)/* (int j = 0; j <= columns.Count; i++)*/
                {
                    oxa = new List<OpenXmlAttribute>
                    {
                        // this is the data type ("t"), with CellValues.String ("str")
                        new OpenXmlAttribute("t", null, column.CellValueType.ToString()),
                        //oxa.Add(new OpenXmlAttribute("s", null, "3"));
                        new OpenXmlAttribute("s", null, column.StyleIndexId ?? "0")
                    };
                    //oxa.Add(new OpenXmlAttribute("s", null, "1"));


                    // it's suggested you also have the cell reference, but
                    // you'll have to calculate the correct cell reference yourself.
                    // Here's an example:
                    //oxa.Add(new OpenXmlAttribute("r", null, "A1"));

                    oxw.WriteStartElement(new Cell(), oxa);
                    {
                        //oxw.WriteElement(new CellValue(string.Format("R{0}C{1}", i, j)));
                        oxw.WriteElement(new CellValue(column.Selector(data[i])));
                    }
                    // this is for Cell
                    oxw.WriteEndElement();
                }

                // this is for Row
                oxw.WriteEndElement();
            }
        }

    }
}
