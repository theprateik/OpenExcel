using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel.Styles
{
    public class OpenExcelBorder : ICloneable
    {
        public OpenExcelBorder(string uid)
        {
            UID = uid;
        }

        public string UID { get; }

        /// <summary>
        /// Creates a duplicate of the current value.
        /// </summary>
        /// <returns>The cloned value.</returns>
        /// <remarks>This method is a deep copy clone.</remarks>
        public object Clone()
        {
            var clone = (OpenExcelBorder)MemberwiseClone();


            return clone;
        }
    }
}
