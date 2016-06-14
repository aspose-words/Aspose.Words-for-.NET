
using System.IO;
using Aspose.Words;
using System;
using Aspose.Words.Markup;
using System.Drawing;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class ComboBoxContentControl
    {
        public static void Run()
        {
            //ExStart:ComboBoxContentControl
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            Document doc = new Document();
            StructuredDocumentTag sdt = new StructuredDocumentTag(doc, SdtType.ComboBox, MarkupLevel.Block);

            sdt.ListItems.Add(new SdtListItem("Choose an item", "-1"));
            sdt.ListItems.Add(new SdtListItem("Item 1", "1"));
            sdt.ListItems.Add(new SdtListItem("Item 2", "2"));
            doc.FirstSection.Body.AppendChild(sdt);

            dataDir = dataDir + "ComboBoxContentControl_out_.docx";
            doc.Save(dataDir);
            //ExEnd:ComboBoxContentControl
            Console.WriteLine("\nCombo box type content control created successfully.\nFile saved at " + dataDir);
        }        
    }
}
