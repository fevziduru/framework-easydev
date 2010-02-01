using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;

namespace EasyDev.Office.Excel
{
    /// <summary>
    /// Extractor factory
    /// </summary>
    public class ExcelExtractorFactory
    {
        public static ExcelExtractor GetExcelExtractor(string filePath)
        {
            return new ExcelExtractor(filePath);
        }
    }
}
