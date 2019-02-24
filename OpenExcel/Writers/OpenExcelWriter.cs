﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcel.Models;
using OpenExcel.Props;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OpenExcel.Writers
{
    public class OpenExcelWriter : IDisposable
    {
        private const uint _rowIdxReset = 0;

        private readonly string _filePath;
        private OpenXmlWriter _workSheetWriter;
        private OpenXmlWriter _workBookWriter;
        private StyleSheetWriter _styleSheetWriter;
        private SharedStringWriter _sharedStringWriter;

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
        private uint _rowIdx;
        private uint _newRowIdx
        {
            get
            {
                _rowIdx++;

                return _rowIdx;
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
            _sharedStringWriter = new SharedStringWriter(_xl);

            _workBookWriter = OpenXmlWriter.Create(_xl.WorkbookPart);
            _workBookWriter.WriteStartElement(new Workbook());
            _workBookWriter.WriteStartElement(new Sheets());
        }

        public void StartCreatingSheet(string sheetName, OpenExcelSheetProperties sheetProperties = default)
        {
            _rowIdx = _rowIdxReset;
            var wsPart = _xl.WorkbookPart.AddNewPart<WorksheetPart>();

            _workBookWriter.WriteElement(new Sheet()
            {
                Name = sheetName,
                SheetId = _newSheetId,
                Id = _xl.WorkbookPart.GetIdOfPart(wsPart)
            });

            _workSheetWriter = OpenXmlWriter.Create(wsPart);
            _workSheetWriter.WriteStartElement(new Worksheet());

            WriteSheetProperties(sheetProperties);
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

            _sharedStringWriter.Close();
            _styleSheetWriter.WriteAndClose();
            _xl.Close();
        }

        public void InsertHeader<T>(List<OpenExcelColumn<T>> columns, int nestedLevel = 0)
        {
            InsertRowToSheet(columns.Select(x => x.Name).ToList(), nestedLevel, CellValues.SharedString);
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
            void insertCells()
            {
                List<OpenXmlAttribute> attributes;

                foreach (var column in columns)
                {
                    var styleIdx = _styleSheetWriter.InsertIfNotExist(column.CellFormat);

                    if (column.CellFormatRule != null)
                    {
                        var cellFormat = column.CellFormatRule(record);
                        styleIdx = _styleSheetWriter.InsertIfNotExist(cellFormat);
                    }


                    string cellValue = column.Selector(record);
                    if (column.CellValueType == CellValues.SharedString)
                    {
                        var sharedStringIdx = _sharedStringWriter.Write(cellValue);

                        cellValue = sharedStringIdx.ToString();
                    }

                    attributes = new List<OpenXmlAttribute>
                    {
                        new OpenXmlAttribute("t", null, column.CellValueType.ToString()), // DataType
                        new OpenXmlAttribute("s", null, styleIdx.ToString()) // Style Index
                    };

                    // it's suggested you also have the cell reference, but
                    // you'll have to calculate the correct cell reference yourself.
                    // Here's an example:
                    //attributes.Add(new OpenXmlAttribute("r", null, "A1"));

                    _workSheetWriter.WriteStartElement(new Cell(), attributes);
                    {
                        _workSheetWriter.WriteElement(new CellValue(cellValue));
                    }
                    _workSheetWriter.WriteEndElement();
                }
            }

            InsertRow(insertCells, nestedLevel);
        }

        public void InsertRowToSheet(List<string> cellValues, int nestedLevel = 0, EnumValue<CellValues> cellValueType = null)
        {
            void insertCells()
            {
                List<OpenXmlAttribute> attributes;

                foreach (var v in cellValues)
                {
                    var cellValue = v;
                    if (cellValueType == CellValues.SharedString)
                    {
                        var sharedStringIdx = _sharedStringWriter.Write(cellValue);
                        cellValue = sharedStringIdx.ToString();
                    }
                    attributes = new List<OpenXmlAttribute>
                    {
                        new OpenXmlAttribute("t", null, cellValueType != null ? cellValueType.ToString() : ((EnumValue<CellValues>)CellValues.String).ToString())
                    };

                    _workSheetWriter.WriteStartElement(new Cell(), attributes);
                    {
                        _workSheetWriter.WriteElement(new CellValue(cellValue));
                    }
                    _workSheetWriter.WriteEndElement();
                }
            }

            InsertRow(insertCells, nestedLevel);
        }

        private void InsertRow(Action insertCells, int nestedLevel)
        {
            List<OpenXmlAttribute> attributes;
            attributes = new List<OpenXmlAttribute>
            {
                new OpenXmlAttribute("r", null, _newRowIdx.ToString())
            };

            if (nestedLevel != 0)
            {
                attributes.Add(new OpenXmlAttribute("outlineLevel", string.Empty, nestedLevel.ToString()));
            }

            _workSheetWriter.WriteStartElement(new Row(), attributes);
            {
                insertCells();
            }
            _workSheetWriter.WriteEndElement();
        }

        private void WriteSheetProperties(OpenExcelSheetProperties sheetProperties)
        {
            if (sheetProperties != null)
            {
                _workSheetWriter.WriteStartElement(new SheetProperties());
                {
                    if (sheetProperties.OutlineProperties != null)
                    {
                        _workSheetWriter.WriteElement(new OutlineProperties
                        {
                            SummaryBelow = sheetProperties.OutlineProperties.SummaryBelow,
                            SummaryRight = sheetProperties.OutlineProperties.SummaryRight
                        });
                    }                    
                }
                _workSheetWriter.WriteEndElement();
            }
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