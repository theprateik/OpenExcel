using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel.Styles
{
    public class OpenExcelFont : ICloneable
    {
        public OpenExcelFont(string uid)
        {
            UID = uid;
        }
        public string UID { get; }
        public uint? FontFamilyNumbering { get; set; }
        public string FontName { get; set; }
        public HexBinaryValue Color { get; set; }
        public uint? FontSize { get; set; }
        public bool Italic { get; set; }
        public bool Bold { get; set; }

        /// <summary>
        /// Creates a duplicate of the current value.
        /// </summary>
        /// <returns>The cloned value.</returns>
        /// <remarks>This method is a deep copy clone.</remarks>
        public object Clone()
        {
            var clone = (OpenExcelFont)MemberwiseClone();
            clone.Color = (HexBinaryValue)Color?.Clone();

            return clone;
        }
    }
}
