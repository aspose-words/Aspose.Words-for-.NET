using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Drawing.Charts;
using Aspose.Words.Fields;
using Aspose.Words.Tables;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class DocumentBuilderInsertElements
    {
        public static void Run()
        {
            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            InsertTextInputFormField(dataDir);
            InsertCheckBoxFormField(dataDir);
            InsertComboBoxFormField(dataDir);
            InsertHtml(dataDir);
            InsertHyperlink(dataDir);
            InsertTableOfContents(dataDir);
            InsertOleObject(dataDir);
        }
        public static void InsertTextInputFormField(string dataDir)
        {
            //ExStart:DocumentBuilderInsertTextInputFormField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertTextInput("TextInput", TextFormFieldType.Regular, "", "Hello", 0);
            dataDir = dataDir + "DocumentBuilderInsertTextInputFormField_out_.doc";
            doc.Save(dataDir);
            //ExEnd:DocumentBuilderInsertTextInputFormField
            Console.WriteLine("\nText input form field using DocumentBuilder inserted successfully into a document.\nFile saved at " + dataDir);
        }
        public static void InsertCheckBoxFormField(string dataDir)
        {
            //ExStart:DocumentBuilderInsertCheckBoxFormField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertCheckBox("CheckBox", true, true, 0);
            dataDir = dataDir + "DocumentBuilderInsertCheckBoxFormField_out_.doc";
            doc.Save(dataDir);
            //ExEnd:DocumentBuilderInsertCheckBoxFormField
            Console.WriteLine("\nCheckbox form field using DocumentBuilder inserted successfully into a document.\nFile saved at " + dataDir);
        }
        public static void InsertComboBoxFormField(string dataDir)
        {
            //ExStart:DocumentBuilderInsertComboBoxFormField
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            string[] items = { "One", "Two", "Three" };
            builder.InsertComboBox("DropDown", items, 0);
            dataDir = dataDir + "DocumentBuilderInsertComboBoxFormField_out_.doc";
            doc.Save(dataDir);
            //ExEnd:DocumentBuilderInsertComboBoxFormField
            Console.WriteLine("\nCombobox form field using DocumentBuilder inserted successfully into a document.\nFile saved at " + dataDir);
        }
        public static void InsertHtml(string dataDir)
        {
            //ExStart:DocumentBuilderInsertHtml
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertHtml(
                "<P align='right'>Paragraph right</P>" +
                "<b>Implicit paragraph left</b>" +
                "<div align='center'>Div center</div>" +
                "<h1 align='left'>Heading 1 left.</h1>");
            dataDir = dataDir + "DocumentBuilderInsertHtml_out_.doc";
            doc.Save(dataDir);
            //ExEnd:DocumentBuilderInsertHtml
            Console.WriteLine("\nHTML using DocumentBuilder inserted successfully into a document.\nFile saved at " + dataDir);
        }
        public static void InsertHyperlink(string dataDir)
        {
            //ExStart:DocumentBuilderInsertHyperlink
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Please make sure to visit ");

            // Specify font formatting for the hyperlink.
            builder.Font.Color = Color.Blue;
            builder.Font.Underline = Underline.Single;
            // Insert the link.
            builder.InsertHyperlink("Aspose Website", "http://www.aspose.com", false);

            // Revert to default formatting.
            builder.Font.ClearFormatting();

            builder.Write(" for more information.");
            dataDir = dataDir + "DocumentBuilderInsertHyperlink_out_.doc";
            doc.Save(dataDir);
            //ExEnd:DocumentBuilderInsertHyperlink
            Console.WriteLine("\nHyperlink using DocumentBuilder inserted successfully into a document.\nFile saved at " + dataDir);
        }
        public static void InsertTableOfContents(string dataDir)
        {
            //ExStart:DocumentBuilderInsertTableOfContents
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a table of contents at the beginning of the document.
            builder.InsertTableOfContents("\\o \"1-3\" \\h \\z \\u");

            // Start the actual document content on the second page.
            builder.InsertBreak(BreakType.PageBreak);

            // Build a document with complex structure by applying different heading styles thus creating TOC entries.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            builder.Writeln("Heading 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

            builder.Writeln("Heading 1.1");
            builder.Writeln("Heading 1.2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            builder.Writeln("Heading 2");
            builder.Writeln("Heading 3");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

            builder.Writeln("Heading 3.1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading3;

            builder.Writeln("Heading 3.1.1");
            builder.Writeln("Heading 3.1.2");
            builder.Writeln("Heading 3.1.3");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

            builder.Writeln("Heading 3.2");
            builder.Writeln("Heading 3.3");

            doc.UpdateFields();
            dataDir = dataDir + "DocumentBuilderInsertTableOfContents_out_.doc";
            doc.Save(dataDir);
            //ExEnd:DocumentBuilderInsertTableOfContents
            Console.WriteLine("\nTable of contents using DocumentBuilder inserted successfully into a document.\nFile saved at " + dataDir);
        }
        public static void InsertOleObject(string dataDir)
        {
            //ExStart:DocumentBuilderInsertOleObject
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertOleObject("http://www.aspose.com", "htmlfile", true, true, null);
            dataDir = dataDir + "DocumentBuilderInsertOleObject_out_.doc";
            doc.Save(dataDir);
            //ExEnd:DocumentBuilderInsertOleObject
            Console.WriteLine("\nOleObject using DocumentBuilder inserted successfully into a document.\nFile saved at " + dataDir);
        }
        
    }
}
