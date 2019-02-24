using OpenExcel.Writers;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel
{
    public static class OpenExcelFactory
    {
        public static OpenExcelWriter CreateOpenExcel(string filePath)
        {
            return new OpenExcelWriter(filePath);
        }
    }
}
