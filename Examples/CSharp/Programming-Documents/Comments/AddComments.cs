using System;
using System.Collections;
using System.IO;

using Aspose.Words;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Comments
{
    class AddComments
    {
        public static void Run()
        {
            //ExStart:AddComments
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithComments();
            //ExStart:CreateSimpleDocumentUsingDocumentBuilder
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Some text is added.");
            //ExEnd:CreateSimpleDocumentUsingDocumentBuilder
            Comment comment = new Comment(doc, "Awais Hafeez", "AH", DateTime.Today);
            builder.CurrentParagraph.AppendChild(comment);
            comment.Paragraphs.Add(new Paragraph(doc));
            comment.FirstParagraph.Runs.Add(new Run(doc, "Comment text."));
           
            dataDir = dataDir + "Comments_out_.doc";
            // Save the document.
            doc.Save(dataDir);
            //ExEnd:AddComments
            Console.WriteLine("\nComments added successfully.\nFile saved at " + dataDir);
        }        
    }
}
