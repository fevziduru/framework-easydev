﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeExcel = Microsoft.Office.Interop.Excel;
using VBControls = Microsoft.Vbe.Interop.Forms;

namespace EasyDev.Office.Excel
{
    public class RadioBoxValueExtractor : ValueExtractor
    {
        #region IValueExtractor Members

        public override object GetValue(OfficeExcel.Workbook workbook, int sheetIndex, string controlName)
        {
            try
            {
                OfficeExcel.OLEObject obj = GetOLEObject(workbook, sheetIndex, controlName);

                if (obj != null)
                {
                    VBControls.OptionButton control = obj.Object as VBControls.OptionButton;
                    if (control != null)
                    {
                        return control.get_Value();
                    }
                    else
                    {
                        return "";
                    }
                }
                else
                {
                    return "";
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public override T GetValue<T>(OfficeExcel.Workbook workbook, int sheetIndex, string controlName)
        {
            return (T)GetValue(workbook, sheetIndex, controlName);
        }

        #endregion
    }
}
