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

                openFolder.Visible = false;
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

            // Check if a word document should be used.
            if (generateLetters.Checked)
            {
                // Generate new filename for the resulting word file.
                string wordLettersFilename = "resultat_" + System.IO.Path.GetFileNameWithoutExtension(shortFileName) + ".docx";
                fileNameInt = 0;

                while (System.IO.File.Exists(
                    System.IO.Path.Combine(directory, wordLettersFilename)))
                {
                    fileNameInt++;
                    wordLettersFilename = string.Format("resultat{0}_{1}",
                        fileNameInt,
                        System.IO.Path.GetFileNameWithoutExtension(shortFileName) + ".docx");
                }
                string wordLettersFullFilename = System.IO.Path.Combine(directory, wordLettersFilename);

                LogOutput("Öppnar ett nytt worddokument för att skriva lite brev...");
                // Generate Word Document
                WordHandler wordFile = new WordHandler();
                wordFile.WordOpenNewDocument();
                
                // Generate the letters to be sent in advance.
                AddHostToWord("Förrätt", participants.FinalStarterHosts, participants.FinalStarterGuests1, participants.FinalStarterGuests2, wordFile);
                AddHostToWord("Huvudrätt", participants.FinalMainCourseHosts, participants.FinalMainCourseGuests1, participants.FinalMainCourseGuests2, wordFile);
                AddHostToWord("Efterrätt", participants.FinalDesertHosts, participants.FinalDesertGuests1, participants.FinalDesertGuests2, wordFile);

                AddInstructionsOfWhereToGoNextToWord("förrätten", participants, wordFile);
                AddInstructionsOfWhereToGoNextToWord("huvudrätten", participants, wordFile);

                // Save the word file.
                wordFile.WordSaveAndClose(System.IO.Path.Combine(wordLettersFullFilename));
                LogOutput("Worddokument skapat: " + wordLettersFullFilename);
            }
            // Activate the open folder button.
            openFolder.Visible = true;

            LogOutput("Klart!");
        }

        /// <summary>
        /// Generates the text that sends the participants to their next meal
        /// </summary>
        /// <param name="meal">"förrätten" or "huvudrätten"</param>
        /// <param name="p">A full class of Foodrelayparticipants</param>
        /// <param name="doc">the word document</param>
        
        private void AddInstructionsOfWhereToGoNextToWord(string meal, FoodRelayParticipants p, WordHandler doc)
        {
            int FindIndex(string needle, List<Participant> haystack)
            {
                int result = -1;
                for (int i = 0; i < haystack.Count(); i++)
                {
                    if(needle== haystack[i].Name)
                    {
                        result = i;
                        break;
                    }
                }
                return result;
            }

            Participant FindInListHaystacks(string needle, List<Participant> Haystack1, List<Participant> Haystack2, List<Participant> HostStack)
            {
                int foundIndex = FindIndex(needle, Haystack1);
                if (foundIndex == -1)
                {
                    foundIndex = FindIndex(needle, Haystack2);
                }
                if(foundIndex != -1)
                {
                    return HostStack[foundIndex];
                }
                else
                {
                    foundIndex = FindIndex(needle, HostStack);
                }   
                if(foundIndex != -1)
                {
                    return null;
                }
                throw new KeyNotFoundException("No new place to go is found for a participant");
            }

            string NextStopSnippet(Participant NextStop)
            {
                char newline = (char)11;
                if (NextStop == null)
                {
                    return "är värd för nästa rätt" + newline + newline;
                }
                else
                {
                    return "är välkomna till" + newline + NextStop.Name + newline + NextStop.ContactInformation + newline + newline;
                }
            }

            List<Participant> hosts = (meal == "förrätten" ? p.FinalStarterHosts : p.FinalMainCourseHosts);
            List<Participant> guests1 = (meal == "förrätten" ? p.FinalStarterGuests1 : p.FinalMainCourseGuests1);
            List<Participant> guests2 = (meal == "förrätten" ? p.FinalStarterGuests2: p.FinalMainCourseGuests2);

            char nl = (char)11; // New-line character in Word.
            for (int i = 0; i < hosts.Count; i++)
            {
                Participant hostNextStop, guest1NextStop, guest2NextStop;
                if (meal == "förrätten")
                {
                    hostNextStop = FindInListHaystacks(hosts[i].Name, p.FinalMainCourseGuests1, p.FinalMainCourseGuests2, p.FinalMainCourseHosts);
                    guest1NextStop = FindInListHaystacks(guests1[i].Name, p.FinalMainCourseGuests1, p.FinalMainCourseGuests2, p.FinalMainCourseHosts);
                    guest2NextStop = FindInListHaystacks(guests2[i].Name, p.FinalMainCourseGuests1, p.FinalMainCourseGuests2, p.FinalMainCourseHosts);
                }
                else
                {
                    hostNextStop = FindInListHaystacks(hosts[i].Name, p.FinalDesertGuests1, p.FinalDesertGuests2, p.FinalDesertHosts);
                    guest1NextStop = FindInListHaystacks(guests1[i].Name, p.FinalDesertGuests1, p.FinalDesertGuests2, p.FinalDesertHosts);
                    guest2NextStop = FindInListHaystacks(guests2[i].Name, p.FinalDesertGuests1, p.FinalDesertGuests2, p.FinalDesertHosts);
                }
                
                string toBeReadBy = "Att läsas under måltiden av:" + nl + hosts[i].Name;
                doc.WordAddText(toBeReadBy, "italic");

                string text =
                    "Hoppas ni har njutit av " + meal + " och sällskapet!" + nl +
                    "Det är fortfarande mycket kvar och nu är det dags att åka vidare..." + nl + nl +
                    hosts[i].Name + nl + NextStopSnippet(hostNextStop) +
                    guests1[i].Name + nl + NextStopSnippet(guest1NextStop) +
                    guests2[i].Name + nl + NextStopSnippet(guest2NextStop);

                doc.WordAddText(text);
                if(!(meal == "huvudrätten" && i == (hosts.Count - 1)))
                {
                    doc.WordAddPageBreak();
                }
            }

        }

        /// <summary>
        /// Generates the text that should be added to the document and sent to the participants in advance.
        /// </summary>
        /// <param name="meal">The part of the meal to prepare</param>
        /// <param name="hosts">list of hosts.</param>
        /// <param name="guests1">first list of guests</param>
        /// <param name="guests2">second list of guests</param>
        /// <param name="doc">the word document</param>
        private void AddHostToWord(string meal, List<Participant> hosts, List<Participant> guests1, List<Participant> guests2, WordHandler doc)
        {
            string getAllergies(string host, string guest1, string guest2)
            {
                string result = "";
                // result += (host == "" ? "" : host);
                result += host;
                if (guest1 != "")
                {
                    if (result == "")
                    {
                        result += guest1;
                    }
                    else
                    {
                        result += ", " + guest1;
                    }
                }
                if (guest2 != "")
                {
                    if (result == "")
                    {
                        result += guest2;
                    }
                    else
                    {
                        result += ", " + guest2;
                    }
                }
                if (result.Count() < 1)
                {
                    result = "Inga allergier";
                }
                
                return result;
            }

            char nl = (char)11; // New-line character in Word.
            for (int i = 0; i < hosts.Count; i++)
            {
                doc.WordAddText(hosts[i].Name, "name");
                string text =
                    "Välkomna att delta i matstafetten!" + nl +
                    "Den del av måltiden ni har blivit tilldelade att förbereda är:" +
                    nl + nl + meal + nl + nl +
                    "Eventuella allergier ni behöver ta hänsyn till är:" + nl;
                string allergies = getAllergies(hosts[i].Allergie, guests1[i].Allergie, guests2[i].Allergie);
                text += allergies;
                doc.WordAddText(text);
                doc.WordAddPageBreak();
            }
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

        /// <summary>
        /// Open the output folder.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OpenFolder_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(@"" + directory);
        }
    }
}
