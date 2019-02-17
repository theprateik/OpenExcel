using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;

namespace OpenExcel
{
    public class ExcelExporter
    {
        public ExcelExporter()
        {

        }

        public void CreateSpreadsheetWorkbook<T>(string filePath, List<T> data, List<OpenExcelColumn<T>> columns)
        {
            using (SpreadsheetDocument xl = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                List<OpenXmlAttribute> oxa;
                OpenXmlWriter oxw;

                xl.AddWorkbookPart();
                WorksheetPart wsp = xl.WorkbookPart.AddNewPart<WorksheetPart>();

                oxw = OpenXmlWriter.Create(wsp);
                {
                    AddStyleSheet(xl);

                    oxw.WriteStartElement(new Worksheet());
                    {
                        oxw.WriteStartElement(new SheetData());
                        {
                            for (int i = 0; i < data.Count; i++)
                            {
                                oxa = new List<OpenXmlAttribute>();
                                // this is the row index
                                //oxa.Add(new OpenXmlAttribute("r", null, i.ToString()));
                                //if (i > 5 && i < 15)
                                //{
                                //    oxa.Add(new OpenXmlAttribute("outlineLevel", string.Empty, "1"));
                                //}
                                oxw.WriteStartElement(new Row(), oxa);

                                foreach(var column in columns)/* (int j = 0; j <= columns.Count; i++)*/
                                {
                                    oxa = new List<OpenXmlAttribute>();
                                    // this is the data type ("t"), with CellValues.String ("str")
                                    oxa.Add(new OpenXmlAttribute("t", null, column.CellValueType.ToString()));
                                    //oxa.Add(new OpenXmlAttribute("s", null, "3"));
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


        public void AddStyleSheet(SpreadsheetDocument document)
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
                        ssWriter.WriteElement(new CellFormat() { FontId = 2U, FillId = 1U, BorderId = 0U });
                        ssWriter.WriteElement(new CellFormat() { NumberFormatId = 164U, FontId = 2U, FillId = 1U, BorderId = 0U, ApplyNumberFormat = true, FormatId = 0, });
                        //ssWriter.WriteElement(new CellFormat() { FontId = 1, FillId = 0 });
                    }
                    ssWriter.WriteEndElement();
                }
                ssWriter.WriteEndElement();
            }
            ssWriter.Close();
        }

        //public static void CreateSpreadsheetWorkbook(string filepath)
        //{
        //    // Create a spreadsheet document by supplying the filepath.
        //    // By default, AutoSave = true, Editable = true, and Type = xlsx.
        //    SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);

        //    // Add a WorkbookPart to the document.
        //    WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();
        //    workbookpart.Workbook = new Workbook();

        //    AddStyleSheet(spreadsheetDocument); // <== Adding stylesheet using above function

        //    // Add a WorksheetPart to the WorkbookPart.
        //    WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
        //    worksheetPart.Worksheet = new Worksheet(new SheetData());

        //    // Add Sheets to the Workbook.
        //    Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

        //    // Append a new worksheet and associate it with the workbook.
        //    Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "mySheet" };
        //    sheets.Append(sheet);

        //    workbookpart.Workbook.Save();

        //    // Close the document.
        //    spreadsheetDocument.Close();
        //}

        //private WorkbookStylesPart AddStyleSheet(SpreadsheetDocument spreadsheet)
        //{
        //    WorkbookStylesPart stylesheet = spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>();

        //    Stylesheet workbookstylesheet = new Stylesheet();

        //    Font font0 = new Font();         // Default font

        //    Font font1 = new Font();         // Bold font
        //    Bold bold = new Bold();
        //    font1.Append(bold);

        //    Fonts fonts = new Fonts();      // <APENDING Fonts>
        //    fonts.Append(font0);
        //    fonts.Append(font1);

        //    // <Fills>
        //    Fill fill0 = new Fill();        // Default fill

        //    Fills fills = new Fills();      // <APENDING Fills>
        //    fills.Append(fill0);

        //    // <Borders>
        //    Border border0 = new Border();     // Defualt border

        //    Borders borders = new Borders();    // <APENDING Borders>
        //    borders.Append(border0);

        //    // <CellFormats>
        //    CellFormat cellformat0 = new CellFormat() { FontId = 0, FillId = 0, BorderId = 0 }; // Default style : Mandatory | Style ID =0

        //    CellFormat cellformat1 = new CellFormat() { FontId = 1 };  // Style with Bold text ; Style ID = 1


        //    // <APENDING CellFormats>
        //    CellFormats cellformats = new CellFormats();
        //    cellformats.Append(cellformat0);
        //    cellformats.Append(cellformat1);


        //    // Append FONTS, FILLS , BORDERS & CellFormats to stylesheet <Preserve the ORDER>
        //    workbookstylesheet.Append(fonts);
        //    workbookstylesheet.Append(fills);
        //    workbookstylesheet.Append(borders);
        //    workbookstylesheet.Append(cellformats);

        //    // Finalize
        //    stylesheet.Stylesheet = workbookstylesheet;
        //    stylesheet.Stylesheet.Save();

        //    return stylesheet;
        //}


    }
}
