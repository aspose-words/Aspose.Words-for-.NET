
using System.IO;
using Aspose.Words;
using System;
using Aspose.Words.Markup;
using System.Drawing;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class RichTextBoxContentControl
    {
        public static void Run()
        {
            //ExStart:RichTextBoxContentControl
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            Document doc = new Document();
            StructuredDocumentTag sdtRichText = new StructuredDocumentTag(doc, SdtType.RichText, MarkupLevel.Block);

            Paragraph para = new Paragraph(doc);
            Run run = new Run(doc);
            run.Text = "Hello World";
            run.Font.Color = Color.Green;
            para.Runs.Add(run);
            sdtRichText.ChildNodes.Add(para);
            doc.FirstSection.Body.AppendChild(sdtRichText);

            dataDir = dataDir + "RichTextBoxContentControl_out_.docx";
            doc.Save(dataDir);
            //ExEnd:RichTextBoxContentControl
            Console.WriteLine("\nRich text box type content control created successfully.\nFile saved at " + dataDir);
        }        
    }
}
