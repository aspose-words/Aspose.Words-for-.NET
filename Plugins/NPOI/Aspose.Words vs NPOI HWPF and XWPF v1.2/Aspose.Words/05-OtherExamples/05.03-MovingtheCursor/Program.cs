using System;
using Aspose.Words;
using System.IO;

namespace _05._03_MovingtheCursor
{
    class Program
    {
        static void Main(string[] args)
        {
            // Check for license and apply if exists
            string licenseFile = AppDomain.CurrentDomain.BaseDirectory + "Aspose.Words.lic";
            if (File.Exists(licenseFile))
            {
                // Apply Aspose.Words API License
                Aspose.Words.License license = new Aspose.Words.License();
                // Place license file in Bin/Debug/ Folder
                license.SetLicense("Aspose.Words.lic");
            }
            Document doc = new Document("../../data/document.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            //Shows how to access the current node in a document builder.
            Node curNode = builder.CurrentNode;
            Paragraph curParagraph = builder.CurrentParagraph;

            // Shows how to move a cursor position to a specified node.
            builder.MoveTo(doc.FirstSection.Body.LastParagraph);

            // Shows how to move a cursor position to the beginning or end of a document.
            builder.MoveToDocumentEnd();
            builder.Writeln("This is the end of the document.");

            builder.MoveToDocumentStart();
            builder.Writeln("This is the beginning of the document.");

            doc.Save("outputDocument.doc");
        }
    }
}
