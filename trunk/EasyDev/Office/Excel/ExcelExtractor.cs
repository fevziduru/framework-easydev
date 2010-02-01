using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.InteropServices;

namespace EasyDev.Office.Excel
{
    /// <summary>
    /// Excel field value extractor
    /// </summary>
    public class ExcelExtractor : IDisposable
    {
        private Microsoft.Office.Interop.Excel.Application excelApp = null;
        private Microsoft.Office.Interop.Excel.Workbook currentWorkbook = null;

        public static readonly global::System.Object Missing = global::System.Type.Missing;        

        private string filePath = string.Empty;

        /// <summary>
        /// Excel file path
        /// </summary>
        public string FilePath { get { return filePath; } }

        /// <summary>
        /// Excel application COM object
        /// </summary>
        public Microsoft.Office.Interop.Excel.Application ExcelApplication
        {
            get
            {
                return excelApp;
            }
        }

        /// <summary>
        /// Default constractor
        /// </summary>
        public ExcelExtractor() { }

        /// <summary>
        /// Custom constractor with filepath parameter
        /// </summary>
        /// <param name="filePath"></param>
        public ExcelExtractor(string filePath)
        {
            Initialize(filePath);
        }

        /// <summary>
        /// Initialize current extractor
        /// </summary>
        /// <param name="filePath"></param>
        protected virtual void Initialize(string filePath)
        {
            try
            {
                if (File.Exists(filePath) == false)
                {
                    throw new IOException(filePath + " is not exist...!");
                }

                excelApp = new Microsoft.Office.Interop.Excel.ApplicationClass();
                currentWorkbook = excelApp.Workbooks.Open(filePath, Missing, false,Missing, Missing, Missing,Missing, Missing, Missing,
                                                      true, Missing, Missing,Missing, Missing, Missing);

            }
            catch (IOException e)
            {
                throw e;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        /// <summary>
        /// Get value by specific type of value extractor.
        /// </summary>
        /// <typeparam name="TExtractor">type of extractor</typeparam>
        /// <typeparam name="TRetValue">return value</typeparam>
        /// <param name="sheetIndex">sheet index</param>
        /// <param name="controlName">control name</param>
        /// <returns>value of the control</returns>
        public TRetValue GetValue<TExtractor, TRetValue>(int sheetIndex, string controlName)
            where TExtractor : IValueExtractor, new()
        {
            return new TExtractor().GetValue<TRetValue>(currentWorkbook, sheetIndex, controlName);
        }

        /// <summary>
        /// Get value by specific type of value extractor.
        /// <remarks>By this method, the first sheet will be used.</remarks>
        /// </summary>
        /// <typeparam name="TExtractor"></typeparam>
        /// <typeparam name="TRetValue"></typeparam>
        /// <param name="controlName"></param>
        /// <returns></returns>
        public TRetValue GetValue<TExtractor, TRetValue>(string controlName)
                    where TExtractor : IValueExtractor, new()
        {
            return new TExtractor().GetValue<TRetValue>(currentWorkbook, 1, controlName);
        }

        #region IDisposable Members

        public void Dispose()
        {
            if (currentWorkbook != null)
            {
                currentWorkbook.Close(false, Missing, false);
                Marshal.FinalReleaseComObject(currentWorkbook);
                currentWorkbook = null;
            }

            if (excelApp != null)
            {
                excelApp.Quit();
                Marshal.FinalReleaseComObject(excelApp);
                excelApp = null;
            }
        }

        #endregion
    }
}
