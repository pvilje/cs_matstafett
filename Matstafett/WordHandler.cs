using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace Matstafett
{
    public class WordHandler
    {
        public Word.Application WordApp { get; set; }
        public Word.Documents WordDocuments { get; set; }
        public Word.Document WordDocument { get; set; }
        public Word.Style WordStyleName { get; set; }
        public Word.Style WordStyleNormalText { get; set; }
        public Word.Style WordStyleItalicText { get; set; }

        public WordHandler()
        {
            WordApp = new Word.Application();
        }

        /// <summary>
        /// Open a new word document.
        /// </summary>
        /// <param name="readOnly">open as readonly or not</param>
        public void WordOpenNewDocument(bool readOnly = false)
        {
            this.WordDocuments = WordApp.Documents;
            this.WordDocument = WordDocuments.Add();

            // Create styles
            WordStyleName = WordDocument.Styles.Add("Namn");
            WordStyleNormalText = WordDocument.Styles.Add("Normal Text");
            WordStyleItalicText = WordDocument.Styles.Add("Italic Text");
            WordStyleName.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
            WordStyleName.Font.Size = 12;
            WordStyleName.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            WordStyleName.Font.Name = "Lucida Calligraphy";

            WordStyleNormalText.Font.Size = 12;
            WordStyleNormalText.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

            WordStyleItalicText.Font.Size = 12;
            WordStyleItalicText.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            WordStyleItalicText.Font.Italic = 1;
        }

        /// <summary>
        /// Save the word document as close Word
        /// </summary>
        /// <param name="filename">the filename</param>
        public void WordSaveAndClose(string filename)
        {
            WordDocument.SaveAs2(filename);
            WordCloseDocument();
        }

        /// <summary>
        /// Closes the word application
        /// </summary>
        public void WordCloseDocument()
        {
            this.WordDocument.Close();
            this.WordApp.Quit();
        }


        public void WordAddText(string text)
        {
            WordDocument.Words.Last.Text = text;
        }

        /// <summary>
        /// Adds a pagebreak last in the document
        /// </summary>
        public void WordAddPageBreak()
        {
            WordDocument.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
        }
    }
}
