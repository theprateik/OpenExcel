using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OpenExcel.Styles;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel
{
    public class StyleSheetWriter : IDisposable
    {
        private readonly SpreadsheetDocument _xl;
        private readonly OpenXmlWriter _writer;
        private readonly Dictionary<string, uint> _cellFormatIdx;
        private readonly Dictionary<string, uint> _fontIdx;
        private readonly Dictionary<string, uint> _numberingFormatIdx;
        private readonly Dictionary<string, uint> _fillIdx;
        private readonly Dictionary<string, uint> _borderIdx;

        private readonly List<OpenExcelFont> _fonts = new List<OpenExcelFont>();
        private readonly List<OpenExcelNumberingFormat> _numberingFormats = new List<OpenExcelNumberingFormat>();
        private readonly List<OpenExcelFill> _fills = new List<OpenExcelFill>();
        private readonly List<OpenExcelBorder> _borders = new List<OpenExcelBorder>();
        private readonly List<CellFormat> _cellFormats = new List<CellFormat>();

        public StyleSheetWriter(SpreadsheetDocument xl)
        {
            _cellFormatIdx = new Dictionary<string, uint>();
            _fontIdx = new Dictionary<string, uint>();
            _numberingFormatIdx = new Dictionary<string, uint>();

            _xl = xl;

            var wbStylesPart = _xl.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            _writer = OpenXmlWriter.Create(wbStylesPart);

            Initialize();
        }

        private void Initialize()
        {
            _fonts.Add(new OpenExcelFont("0"));
            _fontIdx.Add(Guid.NewGuid().ToString(), 1);

            _numberingFormats.Add(new OpenExcelNumberingFormat("0"));
            _numberingFormatIdx.Add(Guid.NewGuid().ToString(), 1);

            _fills.Add(new OpenExcelFill("0"));
            _fillIdx.Add(Guid.NewGuid().ToString(), 1);

            _borders.Add(new OpenExcelBorder("0"));
            _borderIdx.Add(Guid.NewGuid().ToString(), 1);
        }

        private void WriteNumberingFormats()
        {

        }

        public uint InsertIfNotExist(OpenExcelCellFormat cellFormat)
        {
            uint cellFormatIdx;
            if (_cellFormatIdx.TryGetValue(cellFormat.UID, out cellFormatIdx))
            {
                return cellFormatIdx; ;
            }

            uint fontIdx = 0;
            if (cellFormat.Font != null && !_fontIdx.TryGetValue(cellFormat.Font.UID, out fontIdx))
            {
                fontIdx = (uint)(_fontIdx.Count + 1);
                _fontIdx.Add(cellFormat.Font.UID, fontIdx);
            }

            _cellFormatIdx.Add(cellFormat.UID, (uint)(_cellFormatIdx.Count + 1));
            _cellFormats.Add(new CellFormat { FontId = fontIdx, FillId = 0U, BorderId = 0U });
            return (uint)_cellFormats.Count;
        }

        public void Dispose()
        {
            _writer.Dispose();
        }
    }
}
