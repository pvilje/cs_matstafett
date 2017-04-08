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
        public MainWindow()
        {
            InitializeComponent();
        }

        private void browseButton_Click(object sender, EventArgs e)
        {
            // Open a filedialog and save the dialog results.
            OpenFileDialog filePicker = new OpenFileDialog();
            filePicker.Title = "Välj en excelfil";
            filePicker.Filter = "Excel (*.xlsx)|*.xlsx";
            if (filePicker.ShowDialog() == DialogResult.OK)
            {
                // Populate the file name box and hidden file name box
                fileBox.Text = filePicker.FileName;
                hiddenBoxFileName.Text = filePicker.SafeFileName;

                Start.Enabled = true;
            }
        }

        private void start_Click(object sender, EventArgs e)
        {
            if (!generateLetters.Checked == true)
            {
                // ask the user to confirm that letters are not wanted.
                DialogResult noLetters = MessageBox.Show("Inga brev kommer att skapas\nÄr det verkligen meningen?", "Bekräfta att inga brev ska skapas", MessageBoxButtons.YesNo);
                if (noLetters != DialogResult.Yes)
                {
                    return;
                }
            }
            GenerateLineup();
        }

        /* GenerateLineup
         * Main Function that calls all other functions used.
         * Once this function is called everything is autmatic 
         * Params: None */
        private void GenerateLineup()
        {
            // Initiate variables
            FoodRelayParticipants participants = new FoodRelayParticipants();
            string shortFileName = hiddenBoxFileName.Text;
            string fullFileName = fileBox.Text;

            // Get the participant list from the selected excel file. 
            var excelFileParticipantList = new Excel.Application();
            Excel.Workbook wb = excelFileParticipantList.Workbooks.Open(fullFileName, ReadOnly: true);
            Excel.Worksheet ws = wb.Sheets[1];
            for (int i = 1; i <= ws.UsedRange.Rows.Count; i++)
            {
                if (ws.Cells[i, "A"].text != "")
                {
                    participants.AddParticipant(new Participant(
                        name: ws.Cells[i, "A"].text,
                        contact: ws.Cells[i, "B"].text,
                        allergie: ws.Cells[i, "C"].text
                    ));
                }
            }
            wb.Close();
        }
    }
}
