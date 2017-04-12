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
        /// Creates a new Workbook object
        /// </summary>
        public void ExcelCreateSpreadSheet()
        {
            WorkBook = XlApp.Workbooks.Add();
        }

        /// <summary>
        /// Save the document and close it.
        /// </summary>
        /// <param name="filename">the new file name</param>
        public void ExcelSaveAsAndClose(string filename)
        {
            WorkBook.SaveAs(filename);
            ExcelCloseSpreadSheet();
        }

        /// <summary>
        /// Closes the excel spreadsheet
        /// </summary>
        public void ExcelCloseSpreadSheet()
        {
            WorkBook.Close();
            XlApp.Quit();
        }

        /// <summary>
        /// Selects a single WorkSheet, creates new sheets if needed.
        /// </summary>
        /// <param name="sheet">the sheet number</param>
        public void ExcelSelectWorkSheet(int sheet)
        {
            while (WorkBook.Sheets.Count < sheet)
            {
                WorkBook.Sheets.Add(After: WorkBook.Sheets[WorkBook.Sheets.Count]);
            }

            WorkSheet = WorkBook.Sheets[sheet];
        }

        /// <summary>
        /// Adds all the participant's name, contact information and allergies to columns 1,2,3 in the active sheet.
        /// </summary>
        /// <param name="participants">The list of participants to add</param>
        public void AddParticipantList(List<Participant> participants)
        {
            int row = 0;
            foreach (Participant participant in participants)
            {
                row++;
                WorkSheet.Cells[row, 1] = participant.Name;
                WorkSheet.Cells[row, 2] = participant.ContactInformation;
                WorkSheet.Cells[row, 3] = participant.Allergie;
                WorkSheet.Columns.AutoFit();
            } 
        }

        /// <summary>
        /// Adds the complete Lineup to the current sheet.
        /// </summary>
        /// <param name="starterHost">Starter hosts</param>
        /// <param name="starterGuest1">Guest 1 to the starters hosts</param>
        /// <param name="starterGuest2">Guest 2 to the starters hosts</param>
        /// <param name="mainHost">Main Course Hosts</param>
        /// <param name="mainGuest1">Guest 1 to the main course hosts</param>
        /// <param name="mainGuest2">Guest 2 to the main course hosts</param>
        /// <param name="desertHost">Desert hosts</param>
        /// <param name="desertGuest1">Guest 1 to the desert hosts</param>
        /// <param name="desertGuest2">Guest 2 to the desert hosts</param>
        public void AddFoodRelayLineUp(
            List<Participant> starterHost,
            List<Participant> starterGuest1,
            List<Participant> starterGuest2,
            List<Participant> mainHost,
            List<Participant> mainGuest1,
            List<Participant> mainGuest2,
            List<Participant> desertHost,
            List<Participant> desertGuest1,
            List<Participant> desertGuest2)
        {
            // Set up styles.
            Excel.Style h1 = WorkBook.Styles.Add("h1");
            h1.Font.Size = 15;
            h1.Font.Bold = true;
            h1.Font.ColorIndex = 5;  // https://msdn.microsoft.com/en-us/library/cc296089(v=office.12).aspx
            //h1.Borders.ColorIndex = 5;
            //h1.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 2d;

            Excel.Style h1Center = WorkBook.Styles.Add("h1Center");
            h1Center.Font.Size = 15;
            h1Center.Font.Bold = true;
            h1Center.Font.ColorIndex = 5;
            //h1.Borders.ColorIndex = 5;
            //h1.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThick;
            h1Center.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            //(name = "h1", font = Font(size = 15, bold = True, color = "1f497d"),
            //            border = Border(bottom = Side(style = "thick", color = "4f81bd")

            Excel.Range heading = WorkSheet.Cells[1, 1];
            heading.Cells[1,1] = "Sammanfattning";
            heading.Style = h1;

            Excel.Range heading2 = WorkSheet.Cells.Range["C1", "E1"];
            heading2.Cells[1, 1] = "Förrätt";
            heading2.Style = h1Center;
            heading2.MergeCells = true;

            // Add helper functions to add the repeating data.

            WorkSheet.Columns.AutoFit();
        }
    }
}
