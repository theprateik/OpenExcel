using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcel.Styles;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel.Writers
{
    public class StyleSheetWriter : IDisposable
    {
        private const uint _customNumFormatIdStarter = 164U;

        private readonly SpreadsheetDocument _xl;
        private readonly OpenXmlWriter _writer;
        private readonly Dictionary<string, uint> _cellFormatIdx = new Dictionary<string, uint>();
        private readonly Dictionary<string, uint> _fontIdx = new Dictionary<string, uint>();
        private readonly Dictionary<string, uint> _numberingFormatIdx = new Dictionary<string, uint>();
        private readonly Dictionary<string, uint> _fillIdx = new Dictionary<string, uint>();
        private readonly Dictionary<string, uint> _borderIdx = new Dictionary<string, uint>();

        private readonly List<OpenExcelFont> _fonts = new List<OpenExcelFont>();
        private readonly List<OpenExcelNumberingFormat> _numberingFormats = new List<OpenExcelNumberingFormat>();
        private readonly List<OpenExcelFill> _fills = new List<OpenExcelFill>();
        private readonly List<OpenExcelBorder> _borders = new List<OpenExcelBorder>();
        private readonly List<CellFormat> _cellFormats = new List<CellFormat>();

        public StyleSheetWriter(SpreadsheetDocument xl)
        {
            _xl = xl;

            var wbStylesPart = _xl.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            _writer = OpenXmlWriter.Create(wbStylesPart);

            AddInitialStyles();
        }

        private void AddInitialStyles()
        {
            _fonts.Add(new OpenExcelFont("0"));
            _fontIdx.Add(Guid.NewGuid().ToString(), 1);

            _fills.Add(new OpenExcelFill("0") { PatternType = PatternValues.None }); // required, reserved by Excel
            _fillIdx.Add(Guid.NewGuid().ToString(), 0);

            _fills.Add(new OpenExcelFill("1") { PatternType = PatternValues.Gray125 }); // required, reserved by Excel
            _fillIdx.Add(Guid.NewGuid().ToString(), 1);

            _borders.Add(new OpenExcelBorder("0"));
            _borderIdx.Add(Guid.NewGuid().ToString(), 0);

            _cellFormats.Add(new CellFormat { FontId = 0U, FillId = 0U, BorderId = 0U });
            _cellFormatIdx.Add(Guid.NewGuid().ToString(), 0);
        }

        private void WriteNumberingFormats()
        {
            _writer.WriteStartElement(new NumberingFormats());
            {
                foreach(var fmt in _numberingFormats)
                {
                    _writer.WriteElement(new NumberingFormat { NumberFormatId = _customNumFormatIdStarter, FormatCode = fmt.FormatCode });
                }
            }
            _writer.WriteEndElement();
        }

        private void WriteFonts()
        {
            _writer.WriteStartElement(new Fonts());
            {
                foreach(var font in _fonts)
                {
                    _writer.WriteStartElement(new Font());
                    {
                        if (!string.IsNullOrWhiteSpace(font.FontName))
                        {
                            _writer.WriteElement(new FontName { Val = font.FontName });
                        }

                        if (font.FontSize != null)
                        {
                            _writer.WriteElement(new FontSize { Val = font.FontSize });
                        }

                        if (font.Color != null)
                        {
                            _writer.WriteElement(new Color { Rgb = font.Color });
                        }

                        if (font.Italic)
                        {
                            _writer.WriteElement(new Italic());
                        }

                        if (font.Bold)
                        {
                            _writer.WriteElement(new Bold());
                        }
                    }
                    _writer.WriteEndElement();
                }
            }
            _writer.WriteEndElement();
        }

        private void WriteFills()
        {
            _writer.WriteStartElement(new Fills());
            {
                foreach(var fill in _fills)
                {
                    _writer.WriteStartElement(new Fill());
                    {
                        //if (fill.ForegroundColor != null)
                        //{
                        //    var a = (ForegroundColor)fill.ForegroundColor.CloneNode(true);
                        //    _writer.WriteElement(new PatternFill { PatternType = fill.PatternType, ForegroundColor = a, BackgroundColor = fill.BackgroundColor });
                        //}
                        //else
                            _writer.WriteElement(new PatternFill { PatternType = fill.PatternType, ForegroundColor = fill.ForegroundColor, BackgroundColor = fill.BackgroundColor });
                    }
                    _writer.WriteEndElement();
                }
            }
            _writer.WriteEndElement();
        }

        private void WriteBorders()
        {
            _writer.WriteStartElement(new Borders());
            {
                _writer.WriteStartElement(new Border());
                {
                    _writer.WriteElement(new LeftBorder() { Style = BorderStyleValues.DashDot });
                    _writer.WriteElement(new RightBorder() { Style = BorderStyleValues.DashDot });
                    _writer.WriteElement(new TopBorder() { Style = BorderStyleValues.DashDot });
                    _writer.WriteElement(new BottomBorder() { Style = BorderStyleValues.DashDot });
                    _writer.WriteElement(new DiagonalBorder() { Style = BorderStyleValues.DashDot });
                }
                _writer.WriteEndElement();
            }
            _writer.WriteEndElement();
        }

        private void WriteCellFormats()
        {
            _writer.WriteStartElement(new CellFormats());
            {
                foreach (var fmt in _cellFormats)
                {
                    _writer.WriteElement(fmt);
                }
            }
            _writer.WriteEndElement();
        }

        public void WriteAndClose()
        {
            _writer.WriteStartElement(new Stylesheet());
            {
                WriteNumberingFormats();

                WriteFonts();

                WriteFills();

                WriteBorders();

                WriteCellFormats();
            }
            _writer.WriteEndElement();

            _writer.Close();
        }

        public uint InsertIfNotExist(OpenExcelCellFormat cellFormat)
        {
            if (cellFormat == null)
            {
                return 0;
            }

            uint cellFormatIdx;
            if (_cellFormatIdx.TryGetValue(cellFormat.UID, out cellFormatIdx))
            {
                return cellFormatIdx; ;
            }

            uint numFmtIdx = 0;
            if (cellFormat.NumberingFormat != null && !_fontIdx.TryGetValue(cellFormat.NumberingFormat.UID, out numFmtIdx))
            {
                numFmtIdx = (uint)(_numberingFormatIdx.Count) + _customNumFormatIdStarter;
                _numberingFormatIdx.Add(cellFormat.NumberingFormat.UID, numFmtIdx);
                _numberingFormats.Add(cellFormat.NumberingFormat);
            }

            uint fontIdx = 0;
            if (cellFormat.Font != null && !_fontIdx.TryGetValue(cellFormat.Font.UID, out fontIdx))
            {
                fontIdx = (uint)(_fontIdx.Count);
                _fontIdx.Add(cellFormat.Font.UID, fontIdx);
                _fonts.Add(cellFormat.Font);
            }

            uint fillIdx = 0;
            if (cellFormat.Fill != null && !_fillIdx.TryGetValue(cellFormat.Fill.UID, out fillIdx))
            {
                fillIdx = (uint)(_fillIdx.Count);
                _fillIdx.Add(cellFormat.Fill.UID, fillIdx);
                _fills.Add(cellFormat.Fill);
            }

            var newCellFormatIdx = (uint)_cellFormatIdx.Count;
            _cellFormatIdx.Add(cellFormat.UID, (newCellFormatIdx));
            _cellFormats.Add(new CellFormat { NumberFormatId = numFmtIdx, FontId = fontIdx, FillId = fillIdx, BorderId = 0U });
            return newCellFormatIdx;
        }

        public void Dispose()
        {
            _writer.Dispose();
        }
    }
}
