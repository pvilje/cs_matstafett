using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Matstafett
{
    public class ExcelHandler
    {
        public Excel.Application XlApp { get; set; }
        public Excel.Workbook WorkBook { get; set; }
        public Excel.Worksheet WorkSheet { get; set; }

        public ExcelHandler()
        {
            XlApp = new Excel.Application();
        }

        /// <summary>
        /// Opens an excel spredsheet
        /// </summary>
        /// <param name="thisFileName">the filename of the file to open</param>
        /// <param name="readOnly">if read only or not</param>
        public void ExcelOpenSpreadSheet(string thisFileName, bool readOnly = true)
        {
            WorkBook = XlApp.Workbooks.Open(
                thisFileName, 
                ReadOnly: readOnly);
        }

        /// <summary>
        /// Closes the excel spreadsheet
        /// </summary>
        public void ExcelCloseSpreadSheet()
        {
            WorkBook.Close();
        }

        /// <summary>
        /// Selects a single WorkSheet
        /// </summary>
        /// <param name="sheet">the sheet number</param>
        public void ExcelSelectWorkSheet(int sheet)
        {
            WorkSheet = WorkBook.Sheets[sheet];
        }
    }
}
