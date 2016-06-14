
using System.IO;
using Aspose.Words;
using System;
using Aspose.Words.Markup;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class CheckBoxTypeContentControl
    {
        public static void Run()
        {
            //ExStart:CheckBoxTypeContentControl
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            //Open the empty document
            Document doc = new Document();

            DocumentBuilder builder = new DocumentBuilder(doc);
            StructuredDocumentTag SdtCheckBox = new StructuredDocumentTag(doc, SdtType.Checkbox, MarkupLevel.Inline);

            //Insert content control into the document
            builder.InsertNode(SdtCheckBox);
            dataDir = dataDir + "CheckBoxTypeContentControl_out_.docx";

            doc.Save(dataDir, SaveFormat.Docx);
            //ExEnd:CheckBoxTypeContentControl
            Console.WriteLine("\nCheckBox type content control created successfully.\nFile saved at " + dataDir);
        }        
    }
}
