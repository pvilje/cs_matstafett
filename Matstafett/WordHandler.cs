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
            WordStyleName.Font.Italic = 0;
            WordStyleName.Font.Name = "Lucida Calligraphy";

            WordStyleNormalText.Font.Underline = Word.WdUnderline.wdUnderlineNone;
            WordStyleNormalText.Font.Size = 12;
            WordStyleNormalText.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            WordStyleNormalText.Font.Italic = 0;
            WordStyleNormalText.Font.Name = "Calibri";

            WordStyleItalicText.Font.Underline = Word.WdUnderline.wdUnderlineNone;
            WordStyleItalicText.Font.Size = 12;
            WordStyleItalicText.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            WordStyleItalicText.Font.Italic = 1;
            WordStyleItalicText.Font.Name = "Calibri";
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

        /// <summary>
        /// Adds a paragraph to the document with the specified text
        /// </summary>
        /// <param name="text">The string to add</param>
        /// <param name="style">normal, italic or name</param>
        public void WordAddText(string text, string style = "normal")
        {
            Word.Style st = null;
            if (style == "normal")
            {
                st = this.WordStyleNormalText;
            }
            else if (style == "italic")
            {
                st = this.WordStyleItalicText;
            }
            else if (style == "name")
            {
                st = this.WordStyleName;
            }

            // Word.Paragraph p = WordDocument.Words.Last.Paragraphs.Add();
            Word.Paragraph p = WordDocument.Paragraphs.Add();

            p.Range.Text = text;
            p.Range.set_Style(st);
            p.Range.InsertParagraphAfter();
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
