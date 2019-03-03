using System;
using System.Collections.Generic;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenExcel.Styles
{
    public class OpenExcelNumberingFormat : ICloneable
    {
        public OpenExcelNumberingFormat(string uid)
        {
            UID = uid;
        }
        public string UID { get; }
        public string FormatCode { get; set; }

        /// <summary>
        /// Creates a duplicate of the current value.
        /// </summary>
        /// <returns>The cloned value.</returns>
        /// <remarks>This method is a deep copy clone.</remarks>
        public object Clone()
        {
            var clone = (OpenExcelNumberingFormat)MemberwiseClone();


            return clone;
        }
    }
}
