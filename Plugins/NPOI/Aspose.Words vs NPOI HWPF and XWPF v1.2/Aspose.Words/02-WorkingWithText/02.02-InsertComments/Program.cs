using System;
using System.Collections.Generic;
using System.IO;
using System.Text; using Aspose.Words;

namespace _02._02_InsertComments
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
            builder.Write("Some text is added.");

            Comment comment = new Comment(doc, "Aspose", "As", new DateTime());
            builder.CurrentParagraph.AppendChild(comment);
            comment.Paragraphs.Add(new Paragraph(doc));
            comment.FirstParagraph.Runs.Add(new Run(doc, "Comment text."));

            doc.Save("insertedComments.doc");
        }
    }
}
