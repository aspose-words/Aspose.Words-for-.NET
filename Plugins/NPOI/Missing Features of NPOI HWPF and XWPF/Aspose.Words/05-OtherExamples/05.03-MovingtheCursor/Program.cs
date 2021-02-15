using System;
using Aspose.Words;
using System.IO;

namespace _05._03_MovingtheCursor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Check for an Aspose.Words license file in the local file system and apply it, if it exists.
            string licenseFile = AppDomain.CurrentDomain.BaseDirectory + "Aspose.Words.lic";
            if (File.Exists(licenseFile))
            {
                Aspose.Words.License license = new Aspose.Words.License();

                // Use the license from the bin/debug/ Folder.
                license.SetLicense("Aspose.Words.lic");
            }

            Document doc = new Document("../../data/document.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Access the current node in a document builder.
            Node curNode = builder.CurrentNode;
            Paragraph curParagraph = builder.CurrentParagraph;

            // Move the builder's cursor position to a specified node.
            builder.MoveTo(doc.FirstSection.Body.LastParagraph);

            // Move the builder's cursor position to the beginning or end of a document.
            builder.MoveToDocumentEnd();
            builder.Writeln("This is the end of the document.");

            builder.MoveToDocumentStart();
            builder.Writeln("This is the beginning of the document.");

            doc.Save("MovingTheCursor.docx");
        }
    }
}
