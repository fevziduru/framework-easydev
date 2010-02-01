using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeExcel = Microsoft.Office.Interop.Excel;

namespace EasyDev.Office.Excel
{
    public interface IValueExtractor
    {
        object GetValue(OfficeExcel.Workbook workbook, int sheetIndex, string controlName);

        T GetValue<T>(OfficeExcel.Workbook workbook, int sheetIndex, string controlName);
    }
}
