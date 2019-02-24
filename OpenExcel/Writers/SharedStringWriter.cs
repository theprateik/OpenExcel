using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel.Writers
{
    public class SharedStringWriter
    {
        private readonly SpreadsheetDocument _xl;
        private readonly OpenXmlWriter _writer;

        private Dictionary<string, uint> _sharedStringIdx = new Dictionary<string, uint>();
        public SharedStringWriter(SpreadsheetDocument xl)
        {
            _xl = xl;

            var sharedStringTablePart = _xl.WorkbookPart.AddNewPart<SharedStringTablePart>();
            _writer = OpenXmlWriter.Create(sharedStringTablePart);

            Initialize();
        }

        public void Initialize()
        {
            _writer.WriteStartElement(new SharedStringTable());

            // Write initial empty string shared string
            _sharedStringIdx.Add(string.Empty, 0);
            _writer.WriteStartElement(new SharedStringItem());
            {
                _writer.WriteElement(new Text { Text = string.Empty });
            }
            _writer.WriteEndElement();
        }

        public uint Write(string text)
        {
            uint idx = 0;
            if (string.IsNullOrWhiteSpace(text))
            {
                return idx;
            }

            if (_sharedStringIdx.TryGetValue(text, out idx))
            {
                return idx;
            }

            idx = (uint)_sharedStringIdx.Count;
            _sharedStringIdx.Add(text, idx);

            _writer.WriteStartElement(new SharedStringItem());
            {
                _writer.WriteElement(new Text { Text = text });
            }
            _writer.WriteEndElement();

            return idx;
        }

        public void Close()
        {
            _writer.WriteEndElement();
            _writer.Close();
        }
    }
}
