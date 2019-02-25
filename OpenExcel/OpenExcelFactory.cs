using OpenExcel.Apis;
using OpenExcel.Writers;
using System;
using System.Collections.Generic;
using System.Text;

namespace OpenExcel
{
    public static class OpenExcelFactory
    {
        public static OpenExcelApi CreateOpenExcelApi(string filePath)
        {
            return new OpenExcelApi(filePath);
        }
    }
}
