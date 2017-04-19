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
        public Excel.Workbooks WorkBooks { get; set; }
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
            WorkBooks = XlApp.Workbooks;
            WorkBook = WorkBooks.Open(
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
            void addParticipantRange(List<Participant> participants, Excel.Range range)
            {
                int index = 0;
                foreach(Participant participant in participants)
                {
                    index++;
                    range.Cells[index, 1] = participant.Name;
                }
            }

            void addSummaryRange(
                Excel.Style style,
                List<Participant> hosts,
                List<Participant> guests1,
                List<Participant> guests2,
                Excel.Range range)
            {
                range.Cells[1, 1] = "Värd";
                range.Cells[1, 2] = "Gäst 1";
                range.Cells[1, 3] = "Gäst 2";
                range.Range["A1", "C1"].Style = style;
                int index = 1;
                foreach (Participant host in hosts)
                {
                    index++;
                    range.Cells[index, 1] = host.Name;
                    range.Cells[index, 2] = guests1[index - 2].Name;
                    range.Cells[index, 3] = guests2[index - 2].Name;
                }
            }

            // Set up styles.
            Excel.Style h1 = WorkBook.Styles.Add("h1");
            h1.Font.Size = 15;
            h1.Font.Bold = true;
            h1.Font.ColorIndex = 5;  // https://msdn.microsoft.com/en-us/library/cc296089(v=office.12).aspx

            Excel.Style h1Center = WorkBook.Styles.Add("h1Center");
            h1Center.Font.Size = 15;
            h1Center.Font.Bold = true;
            h1Center.Font.ColorIndex = 5;
            h1Center.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            Excel.Style h2 = WorkBook.Styles.Add("h2");
            h2.Font.Size = 13;
            h2.Font.Bold = true;
            h2.Font.ColorIndex = 32;

            Excel.Style h2center = WorkBook.Styles.Add("h2Center");
            h2center.Font.Size = 13;
            h2center.Font.Bold = true;
            h2center.Font.ColorIndex = 32;
            h2center.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            // Add summary to column 1
            // Add detailed content to column 3-5
            // ==================================

            // Headers
            Excel.Range heading = WorkSheet.Cells[1, 1];
            heading.Cells[1, 1] = "Sammanfattning";
            heading.Style = h1;

            // Starters
            Excel.Range starterHeader = WorkSheet.Cells[2, 1];
            starterHeader.Cells[1, 1] = "Värd Förrätt:";
            starterHeader.Style = h2;

            addParticipantRange(
                starterHost, WorkSheet.Cells.Range[
                    string.Format("A3"),
                    string.Format("A{0}",starterHost.Count + 3)
                    ]
                );

            Excel.Range headingStarterMerged = WorkSheet.Cells.Range["C1", "E1"];
            headingStarterMerged.Cells[1, 1] = "Förrätt";
            headingStarterMerged.Style = h1Center;
            headingStarterMerged.MergeCells = true;

            addSummaryRange(
                h2center,
                starterHost, 
                starterGuest1,
                starterGuest2,
                WorkSheet.Cells.Range[
                    string.Format("C2"),
                    string.Format("E{0}", starterHost.Count + 3)
                    ]
                );

            // Main Course
            Excel.Range mainHeader = WorkSheet.Cells[mainHost.Count + 5, 1];
            mainHeader.Cells[1, 1] = "Värd Huvudrätt:";
            mainHeader.Style = h2;
 
            addParticipantRange(mainHost, WorkSheet.Cells.Range[
                string.Format("A{0}", mainHost.Count + 6),
                string.Format("A{0}", mainHost.Count * 2 + 6)
                ]);

            Excel.Range headingMainMerged = WorkSheet.Cells.Range[
                string.Format("C{0}",mainHost.Count + 4), 
                string.Format("E{0}", mainHost.Count + 4)
                ];
            headingMainMerged.Cells[1, 1] = "Huvudrätt";
            headingMainMerged.Style = h1Center;
            headingMainMerged.MergeCells = true;

            addSummaryRange(
                h2center,
                mainHost,
                mainGuest1,
                mainGuest2,
                WorkSheet.Cells.Range[
                    string.Format("C{0}", mainHost.Count + 5),
                    string.Format("E{0}", mainHost.Count * 2 + 5)
                    ]
                );

            // Desert
            Excel.Range desertHeader = WorkSheet.Cells[desertHost.Count*2 + 8, 1];
            desertHeader.Cells[1, 1] = "Värd Efterrätt:";
            desertHeader.Style = h2;

            addParticipantRange(desertHost, WorkSheet.Cells.Range[
                string.Format("A{0}", desertHost.Count * 2 + 9),
                string.Format("A{0}", desertHost.Count * 3 + 9)
                ]);

            Excel.Range headingDesertMerged = WorkSheet.Cells.Range[
                string.Format("C{0}", desertHost.Count * 2 + 7),
                string.Format("E{0}", desertHost.Count * 2 + 7)
                ];
            headingDesertMerged.Cells[1, 1] = "Huvudrätt";
            headingDesertMerged.Style = h1Center;
            headingDesertMerged.MergeCells = true;

            addSummaryRange(
                h2center,
                desertHost,
                desertGuest1,
                desertGuest2,
                WorkSheet.Cells.Range[
                    string.Format("C{0}", desertHost.Count * 2 + 8),
                    string.Format("E{0}", desertHost.Count * 3 + 8)
                    ]
                );

            WorkSheet.Columns.AutoFit();
        }
    }
}
