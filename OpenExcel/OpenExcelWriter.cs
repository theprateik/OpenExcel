using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OpenExcel
{
    public class OpenExcelWriter : IDisposable
    {
        private readonly string _filePath;
        private OpenXmlWriter _workSheetWriter;
        private OpenXmlWriter _workBookWriter;
        private StyleSheetWriter _styleSheetWriter;

        private readonly SpreadsheetDocument _xl;
        private uint _sheetId = 0;
        private uint _newSheetId
        {
            get
            {
                _sheetId++;
                return _sheetId;
            }
        }
        public OpenExcelWriter(string filePath)
        {
            _filePath = filePath;
            _xl = SpreadsheetDocument.Create(_filePath, SpreadsheetDocumentType.Workbook);

            Initialize();
        }

        private void Initialize()
        {
            _xl.AddWorkbookPart();

            _styleSheetWriter = new StyleSheetWriter(_xl);

            //WriteStyleSheet();

            _workBookWriter = OpenXmlWriter.Create(_xl.WorkbookPart);
            _workBookWriter.WriteStartElement(new Workbook());
            _workBookWriter.WriteStartElement(new Sheets());
        }

        public void StartCreatingSheet(string sheetName)
        {
            var wsPart = _xl.WorkbookPart.AddNewPart<WorksheetPart>();

            _workBookWriter.WriteElement(new Sheet()
            {
                Name = sheetName,
                SheetId = _newSheetId,
                Id = _xl.WorkbookPart.GetIdOfPart(wsPart)
            });

            _workSheetWriter = OpenXmlWriter.Create(wsPart);
            _workSheetWriter.WriteStartElement(new Worksheet());
            _workSheetWriter.WriteStartElement(new SheetData());
        }

        public void EndCreatingSheet()
        {
            _workSheetWriter.WriteEndElement(); // End Writing SheetData
            _workSheetWriter.WriteEndElement(); // End Writing Worksheet
            _workSheetWriter.Close();
        }

        public void EndCreatingWorkbook()
        {
            _workBookWriter.WriteEndElement();  // End Writing Sheets
            _workBookWriter.WriteEndElement(); // End Writing Workbook 
            _workBookWriter.Close();

            //WriteStyleSheet();
            _styleSheetWriter.Write();

            _xl.Close();
        }

        private void WriteStyleSheet()
        {
            var ss = _xl.WorkbookPart.AddNewPart<WorkbookStylesPart>();

            using (var ssWriter = OpenXmlWriter.Create(ss))
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
                            ssWriter.WriteElement(new Italic());
                            ssWriter.WriteElement(new FontSize() { Val = 11 });
                        }
                        ssWriter.WriteEndElement();

                        ssWriter.WriteStartElement(new Font());
                        {
                            ssWriter.WriteElement(new Bold());
                            ssWriter.WriteElement(new FontSize() { Val = 12 });
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
                        ssWriter.WriteElement(new CellFormat() { FontId = 1U, FillId = 0U, BorderId = 0U });
                        ssWriter.WriteElement(new CellFormat() { FontId = 2U, FillId = 0U, BorderId = 0U });
                        ssWriter.WriteElement(new CellFormat() { NumberFormatId = 164U, FontId = 2U, FillId = 1U, BorderId = 0U, ApplyNumberFormat = true, FormatId = 0, });
                        //ssWriter.WriteElement(new CellFormat() { FontId = 1, FillId = 0 });
                    }
                    ssWriter.WriteEndElement();
                }
                ssWriter.WriteEndElement();

                ssWriter.Close();
            }
        }

        public void InsertHeader<T>(List<OpenExcelColumn<T>> columns, int nestedLevel = 0)
        {
            InsertRowToSheet(columns.Select(x => x.Name).ToList(), nestedLevel);
        }

        public void InsertDataSetToSheet<T>(List<T> data, List<OpenExcelColumn<T>> columns, int nestedLevel = 0)
        {
            for (int i = 0; i < data.Count; i++)
            {
                InsertRowToSheet(data[i], columns, nestedLevel);
            }
        }

        public void InsertRowToSheet<T>(T record, List<OpenExcelColumn<T>> columns, int nestedLevel = 0)
        {
            List<OpenXmlAttribute> attributes;
            //rowCounter++;
            attributes = new List<OpenXmlAttribute>();
            // this is the row index
            //attributes.Add(new OpenXmlAttribute("r", null, i.ToString()));
            if (nestedLevel != 0)
            {
                attributes.Add(new OpenXmlAttribute("outlineLevel", string.Empty, nestedLevel.ToString()));
            }
            _workSheetWriter.WriteStartElement(new Row(), attributes);

            foreach (var column in columns)/* (int j = 0; j <= columns.Count; i++)*/
            {
                var styleIdx = _styleSheetWriter.InsertIfNotExist(column.CellFormat);
                if (column.CellFormatRule != null)
                {
                    var cellFormat = column.CellFormatRule(record);
                    styleIdx = _styleSheetWriter.InsertIfNotExist(cellFormat);
                }

                attributes = new List<OpenXmlAttribute>
                    {
                        // this is the data type ("t"), with CellValues.String ("str")
                        new OpenXmlAttribute("t", null, column.CellValueType.ToString()),
                        //attributes.Add(new OpenXmlAttribute("s", null, "3"));
                        new OpenXmlAttribute("s", null, styleIdx.ToString())
                        //new OpenXmlAttribute("s", null, column.StyleIndexId)
                    };
                //attributes.Add(new OpenXmlAttribute("s", null, "1"));


                // it's suggested you also have the cell reference, but
                // you'll have to calculate the correct cell reference yourself.
                // Here's an example:
                //attributes.Add(new OpenXmlAttribute("r", null, "A1"));

                _workSheetWriter.WriteStartElement(new Cell(), attributes);
                {
                    //_writer.WriteElement(new CellValue(string.Format("R{0}C{1}", i, j)));
                    _workSheetWriter.WriteElement(new CellValue(column.Selector(record)));
                }
                // this is for Cell
                _workSheetWriter.WriteEndElement();
            }

            // this is for Row
            _workSheetWriter.WriteEndElement();
        }

        public void InsertRowToSheet(List<string> cellValues, int nestedLevel = 0)
        {
            List<OpenXmlAttribute> attributes;
            attributes = new List<OpenXmlAttribute>();
            if (nestedLevel != 0)
            {
                attributes.Add(new OpenXmlAttribute("outlineLevel", string.Empty, nestedLevel.ToString()));
            }
            _workSheetWriter.WriteStartElement(new Row(), attributes);
            {
                foreach(var v in cellValues)
                {
                    attributes = new List<OpenXmlAttribute>
                    {
                        new OpenXmlAttribute("t", null, "str")
                    };

                    _workSheetWriter.WriteStartElement(new Cell(), attributes);
                    {
                        //_writer.WriteElement(new CellValue(string.Format("R{0}C{1}", i, j)));
                        _workSheetWriter.WriteElement(new CellValue(v));
                    }
                    _workSheetWriter.WriteEndElement();
                }
            }
            _workSheetWriter.WriteEndElement();

        }

        public void Dispose()
        {
            _workSheetWriter.Dispose();
            _workBookWriter.Dispose();
            _styleSheetWriter.Dispose();
            _xl.Dispose();
        }
    }
}
