using Aspose.Words;
using Aspose.Words.Tables;

namespace QA_Tests
{
    /// <summary>
    /// Functions for operations with document and content
    /// </summary>
    internal static class DocumentHelper
    {
        /// <summary>
        /// Create new document without run in the paragraph
        /// </summary>
        internal static Document CreateDocumentWithoutDummyText()
        {
            Document doc = new Document();

            //Remove the previous changes of the document
            doc.RemoveAllChildren();

            //Set the document author
            doc.BuiltInDocumentProperties.Author = "Test Author";

            //Create paragraph without run
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln();

            return doc;
        }

        /// <summary>
        /// Create new document with text
        /// </summary>
        internal static Document CreateDocumentFillWithDummyText()
        {
            Document doc = new Document();

            //Remove the previous changes of the document
            doc.RemoveAllChildren();

            //Set the document author
            doc.BuiltInDocumentProperties.Author = "Test Author";

            DocumentBuilder builder = new DocumentBuilder(doc);

            //Insert new table with two rows and two cells
            InsertTable(doc);

            //Insert new paragraph with text
            builder.Writeln("Hello World!");

            return doc;
        }


        /// <summary>
        /// Insert new table in the document
        /// </summary>
        private static void InsertTable(Document doc)
        {
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            //Start creating a new table
            Table table = builder.StartTable();

            //Insert Row 1 Cell 1
            builder.InsertCell();
            builder.Write("Date");

            //Set width to fit the table contents
            table.AutoFit(AutoFitBehavior.AutoFitToContents);
            
            //Insert Row 1 Cell 2
            builder.InsertCell();
            builder.Write(" ");

            builder.EndRow();

            //Insert Row 2 Cell 1
            builder.InsertCell();
            builder.Write("Author");

            //Insert Row 2 Cell 2
            builder.InsertCell();
            builder.Write(" ");

            builder.EndRow();

            builder.EndTable();
        }

        /// <summary>
        /// Insert text into the current document
        /// </summary>
        /// <param name="doc">
        /// Current document
        /// </param>
        /// <param name="text">
        /// Custom text
        /// </param>
        internal static Run InsertNewRun(Document doc, string text)
        {
            Paragraph para = GetParagraph(doc, 0);

            Run run = new Run(doc) { Text = text };

            para.AppendChild(run);

            return run;
        }

        /// <summary>
        /// Get paragraph text of the current document
        /// </summary>
        /// <param name="doc">
        /// Current document
        /// </param>
        /// <param name="paraIndex">
        /// Paragraph number from collection
        /// </param>
        internal static string GetParagraphText(Document doc, int paraIndex)
        {
            return doc.FirstSection.Body.Paragraphs[paraIndex].GetText();
        }

        /// <summary>
        /// Get paragraph of the current document
        /// </summary>
        /// <param name="doc">
        /// Current document
        /// </param>
        /// <param name="paraIndex">
        /// Paragraph number from collection
        /// </param>
        internal static Paragraph GetParagraph(Document doc, int paraIndex)
        {
            return doc.FirstSection.Body.Paragraphs[paraIndex];
        }
    }
}
