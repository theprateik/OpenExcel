using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel.Styles
{
    public class OpenExcelFill : ICloneable
    {
        public OpenExcelFill(string uid)
        {
            UID = uid;
        }

        public string UID { get; }

        public EnumValue<PatternValues> PatternType { get; set; }

        public ForegroundColor ForegroundColor { get; set; }

        public BackgroundColor BackgroundColor { get; set; }

        /// <summary>
        /// Creates a duplicate of the current value.
        /// </summary>
        /// <returns>The cloned value.</returns>
        /// <remarks>This method is a deep copy clone.</remarks>
        public object Clone()
        {
            var clone = (OpenExcelFill)MemberwiseClone();
            clone.PatternType = (EnumValue<PatternValues>)PatternType?.Clone();
            clone.ForegroundColor = (ForegroundColor)ForegroundColor?.Clone();
            clone.BackgroundColor = (BackgroundColor)BackgroundColor?.Clone();

            return clone;
        }
    }
}
