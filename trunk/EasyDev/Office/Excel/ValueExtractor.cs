using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeExcel = Microsoft.Office.Interop.Excel;

namespace EasyDev.Office.Excel
{
    public abstract class ValueExtractor : IValueExtractor
    {
        protected virtual OfficeExcel.OLEObject GetOLEObject(OfficeExcel.Workbook workbook, int sheetIndex, string objectName)
        {
            OfficeExcel.Worksheet sheet = workbook.Worksheets[sheetIndex] as OfficeExcel.Worksheet;
            if (sheet != null)
            {
                return sheet.OLEObjects(objectName) as OfficeExcel.OLEObject;
            }

            return null;
        }

        #region IValueExtractor Members

        public virtual object GetValue(Microsoft.Office.Interop.Excel.Workbook workbook, int sheetIndex, string controlName)
        {
            throw new NotSupportedException();
        }

        public virtual T GetValue<T>(Microsoft.Office.Interop.Excel.Workbook workbook, int sheetIndex, string controlName)
        {
            throw new NotSupportedException();
        }

        #endregion
    }
}
