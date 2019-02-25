using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcel.Models;
using OpenExcel.Props;
using OpenExcel.Writers;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OpenExcel.Apis
{
    public class OpenExcelApi : IDisposable
    {
        private const uint _rowIdxReset = 0;

        private readonly string _filePath;
        private readonly SpreadsheetDocument _xl;
        private readonly StyleSheetWriter _styleSheetWriter;
        private readonly SharedStringWriter _sharedStringWriter;

        private OpenXmlWriter _workSheetWriter;
        private OpenXmlWriter _workBookWriter;
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

        public OpenExcelApi(string filePath)
        {
            _filePath = filePath;
            _xl = SpreadsheetDocument.Create(_filePath, SpreadsheetDocumentType.Workbook);

            _xl.AddWorkbookPart();

            _styleSheetWriter = new StyleSheetWriter(_xl);
            _sharedStringWriter = new SharedStringWriter(_xl);

            Initialize();
        }

        private void Initialize()
        {
            _workBookWriter = OpenXmlWriter.Create(_xl.WorkbookPart);
            _workBookWriter.WriteStartElement(new Workbook());
            _workBookWriter.WriteStartElement(new Sheets());
        }

        public void WriteStartSheet(string sheetName, OpenExcelSheetProperties sheetProperties = default)
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

        public void WriteEndSheet()
        {
            _workSheetWriter.WriteEndElement(); // End Writing SheetData
            _workSheetWriter.WriteEndElement(); // End Writing Worksheet
            _workSheetWriter.Close();
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

        public void Close()
        {
            _workBookWriter.WriteEndElement();  // End Writing Sheets
            _workBookWriter.WriteEndElement(); // End Writing Workbook 

            _workBookWriter.Close();

            _sharedStringWriter.Close();
            _styleSheetWriter.WriteAndClose();

            _xl.Close();
        }

        public uint WriteStartRow(OpenExcelRowProperties rowProperties)
        {
            List<OpenXmlAttribute> attributes;

            var rowNum = _newRowIdx;
            attributes = new List<OpenXmlAttribute>
            {
                new OpenXmlAttribute("r", null, rowNum.ToString())
            };

            if (rowProperties != null && rowProperties.OutlineLevel != 0)
            {
                attributes.Add(new OpenXmlAttribute("outlineLevel", string.Empty, rowProperties.OutlineLevel.ToString()));
            }

            _workSheetWriter.WriteStartElement(new Row(), attributes);

            return rowNum;
        }

        public void WriteEndRow()
        {
            _workSheetWriter.WriteEndElement();
        }

        public void WriteCell(string cellValue, OpenExcelCellProperties cellProperties)
        {
            cellProperties = cellProperties ?? new OpenExcelCellProperties();

            List<OpenXmlAttribute> attributes;
            if (cellProperties.DataType == CellValues.SharedString)
            {
                var sharedStringIdx = _sharedStringWriter.Write(cellValue);
                cellValue = sharedStringIdx.ToString();
            }
            attributes = new List<OpenXmlAttribute>
            {
                new OpenXmlAttribute("t", null, cellProperties.DataType != null ? cellProperties.DataType.ToString() : ((EnumValue<CellValues>)CellValues.String).ToString()), // DataType
                new OpenXmlAttribute("s", null, cellProperties.StyleIdx.ToString()) // Style Index
            };

            _workSheetWriter.WriteStartElement(new Cell(), attributes);
            {
                _workSheetWriter.WriteElement(new CellValue(cellValue));
            }
            _workSheetWriter.WriteEndElement();
        }

        public void WriteRow(List<string> cellValues, OpenExcelRowProperties rowProperties = default, EnumValue<CellValues> cellValueType = null)
        {
            WriteStartRow(rowProperties);

            foreach (var v in cellValues)
            {
                WriteCell(v, new OpenExcelCellProperties { DataType = cellValueType });
            }
            
            WriteEndRow();
        }

        public void WriteRow<T>(T record, List<OpenExcelColumn<T>> columns, OpenExcelRowProperties rowProperties = default)
        {
            WriteStartRow(rowProperties);

            for (var i = 0; i < columns.Count; i++)
            {
                var styleIdx = _styleSheetWriter.InsertIfNotExist(columns[i].CellFormat);

                if (columns[i].CellFormatRule != null)
                {
                    var cellFormat = columns[i].CellFormatRule(record, _rowIdx, (uint)(i+1));
                    styleIdx = _styleSheetWriter.InsertIfNotExist(cellFormat);
                }

                string cellValue = columns[i].Selector(record);

                WriteCell(cellValue, new OpenExcelCellProperties { DataType = columns[i].CellValueType, StyleIdx = styleIdx });
            }

            WriteEndRow();
        }

        public void WriteRowSet<T>(List<T> data, List<OpenExcelColumn<T>> columns, OpenExcelRowProperties rowProperties = default)
        {
            for (int i = 0; i < data.Count; i++)
            {
                WriteRow(data[i], columns, rowProperties);
            }
        }

        public void Dispose()
        {
            _workSheetWriter.Dispose();
            _workBookWriter.Dispose();
            _styleSheetWriter.Dispose();
            _sharedStringWriter.Dispose();

            _xl.Dispose();
        }
    }
}
