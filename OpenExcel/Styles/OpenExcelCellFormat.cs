using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel.Styles
{
    public class OpenExcelCellFormat : ICloneable
    {
        public OpenExcelCellFormat(string uid)
        {
            UID = uid;
        }

        public string UID { get; }

        public OpenExcelNumberingFormat NumberingFormat { get; set; }

        public OpenExcelFont Font { get; set; }

        public OpenExcelFill Fill { get; set; }

        public OpenExcelBorder Border { get; set; }

        /// <summary>
        /// Creates a duplicate of the current value.
        /// </summary>
        /// <returns>The cloned value.</returns>
        /// <remarks>This method is a deep copy clone.</remarks>
        public object Clone()
        {
            var clone = (OpenExcelCellFormat)MemberwiseClone();
            clone.NumberingFormat = (OpenExcelNumberingFormat)NumberingFormat?.Clone();
            clone.Font = (OpenExcelFont)Font?.Clone();
            clone.Fill = (OpenExcelFill)Fill?.Clone();
            clone.Border = (OpenExcelBorder)Border?.Clone();

            return clone;
        }
    }
}
