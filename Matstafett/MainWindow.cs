using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace Matstafett
{
    public partial class MainWindow : Form
    {
        public string shortFileName;
        public string fullFileName;
        public string directory;

        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// OnClick handler for the browsebutton
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BrowseButton_Click(object sender, EventArgs e)
        { 
            // Open a filedialog and save the dialog results.
            OpenFileDialog filePicker = new OpenFileDialog();
            filePicker.Title = "Välj en excelfil";
            filePicker.Filter = "Excel (*.xlsx)|*.xlsx";
            if (filePicker.ShowDialog() == DialogResult.OK)
            {
                // Populate the file name box and save the file details.
                fullFileName = filePicker.FileName;
                shortFileName = filePicker.SafeFileName;
                directory = System.IO.Path.GetDirectoryName(filePicker.FileName);
                fileBox.Text = fullFileName;

                Start.Enabled = true;
                LogOutput(string.Format("Vald fil: {0}", filePicker.SafeFileName));
            }
        }

        /// <summary>
        /// OnClick handler for the Clear log button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ClearLog_Click(object sender, EventArgs e)
        {
            log.Clear();
        }

        /// <summary>
        /// OnClick Handler for the start Button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Start_Click(object sender, EventArgs e)
        {
            if (!generateLetters.Checked == true)
            {
                // ask the user to confirm that letters are not wanted.
                DialogResult noLetters = MessageBox.Show("Inga brev kommer att skapas\nÄr det verkligen meningen?", "Bekräfta att inga brev ska skapas", MessageBoxButtons.YesNo);
                if (noLetters != DialogResult.Yes)
                {
                    LogOutput("Användaren ångrade sig visst... Avvaktar vidare instruktioner");
                    return;
                }
            }
            LogOutput("Kul! Då kör vi!");
            GenerateLineup();
        }

         /// <summary>
         /// Main Function that calls all other functions used.
         /// Once this function is called everything is generate autmatically.
         /// So should not be called until all options are handled.
         /// </summary>
        private void GenerateLineup()
        {
            // Initiate variables.
            // *******************
            FoodRelayParticipants participants = new FoodRelayParticipants();

            // Get the participant list from the selected excel file. 
            // ******************************************************
            LogOutput("Letar efter deltagare i filen.");

            ExcelHandler excelFileParticipants = new ExcelHandler();
            excelFileParticipants.ExcelOpenSpreadSheet(fullFileName);
            excelFileParticipants.ExcelSelectWorkSheet(1);
            for (int i = 1; i <= excelFileParticipants.WorkSheet.UsedRange.Rows.Count; i++)
            {
                if (excelFileParticipants.WorkSheet.Cells[i, "A"].text != "")
                {
                    participants.AddParticipant(new Participant(
                        name: excelFileParticipants.WorkSheet.Cells[i, "A"].text,
                        contact: excelFileParticipants.WorkSheet.Cells[i, "B"].text,
                        allergie: excelFileParticipants.WorkSheet.Cells[i, "C"].text
                    ));
                }
            }
            excelFileParticipants.ExcelCloseSpreadSheet();
            LogOutput(string.Format("Hittade {0} deltagare.", participants.All.Count));

            // Verify the number of Participants
            // *********************************
            LogOutput("Försäkrar mig om att antalet deltagare är ok.");
            int numberOfParticipantsOk = participants.ValidateNumberOfParticipants();
            if (numberOfParticipantsOk == 1)
            {
                LogOutput("Glöm det, avbryter... Det MÅSTE vara fler än 9 deltagare! Avbryter.");
                return;
            }
            else if (numberOfParticipantsOk == 2)
            {
                LogOutput("Nejje! Antalet Deltagare måste vara delbart med tre. Avbryter.");
                return;
            }

            // Randomize array index. 
            LogOutput("Skapar en slumpad lista... Faktiskt lika lång som antalet deltagare :)");
            participants.GenerateRandomizedIndex();

            // Place the participants into groups.
            LogOutput("Använder min slumpade lista för att skyffla runt deltagarna");
            participants.PlaceParticipantsIntoGroups();

            // Create the final Lineup
            LogOutput("Genererar den slutgiltiga uppställningen");
            participants.GenerateLineup();

            // Generate new filename for the resulting excel file.
            string excelResultFileName = "resultat_" + shortFileName;
            int fileNameInt = 0;
            
            while (System.IO.File.Exists(
                System.IO.Path.Combine(directory, excelResultFileName)))
            {
                fileNameInt++;
                excelResultFileName = string.Format("resultat{0}_{1}",
                    fileNameInt,
                    shortFileName);
            }
            string excelResultFullFileName = System.IO.Path.Combine(directory, excelResultFileName);

            // Open a new Excel WorkBook
            LogOutput("Öppnar en ny excelfil och sparar resultatet:" + excelResultFileName);
            ExcelHandler excelResultFile = new ExcelHandler();
            excelResultFile.ExcelCreateSpreadSheet();

            // Write all participants to sheet 1
            excelResultFile.ExcelSelectWorkSheet(1);
            excelResultFile.WorkSheet.Name = "Deltagarlista";
            excelResultFile.AddParticipantList(participants.AllSorted);

            // Write a nice summary to sheet 2
            excelResultFile.ExcelSelectWorkSheet(2);
            excelResultFile.WorkSheet.Name = "Matstafett uppställning";
            excelResultFile.AddFoodRelayLineUp(
                participants.FinalStarterHosts,
                participants.FinalStarterGuests1,
                participants.FinalStarterGuests2,
                participants.FinalMainCourseHosts,
                participants.FinalMainCourseGuests1,
                participants.FinalMainCourseGuests2,
                participants.FinalDesertHosts,
                participants.FinalDesertGuests1,
                participants.FinalDesertGuests2);

            // Save the excel file
            excelResultFile.ExcelSaveAsAndClose(excelResultFullFileName);
            LogOutput("Excelfil skapad: " + excelResultFullFileName);

        }

        /// <summary>
        /// Short function to log output to the log textbox
        /// </summary>
        /// <param name="text">The string to add to the log</param>
        private void LogOutput(string text)
        {
            log.AppendText(text + Environment.NewLine);
        }

        /// <summary>
        /// OnClick handler for the menu item Instruktioner
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void InstruktionerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new Instructions().ShowDialog();
        }

        /// <summary>
        /// OnClick handler for the Menu "Krav på Excelfilen"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void KravPåFilenToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new ExcelFileRequirements().ShowDialog();
        }

        /// <summary>
        /// Show the About boc
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            new AboutBox().ShowDialog();
        }
    }
}
