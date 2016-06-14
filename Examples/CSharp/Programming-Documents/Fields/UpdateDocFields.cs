using System;
using System.Collections;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class UpdateDocFields
    {
        public static void Run()
        {
            //ExStart:UpdateDocFields
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();

            Document doc = new Document(dataDir + "Rendering.doc");

            // This updates all fields in the document.
            doc.UpdateFields();
            dataDir = dataDir + "Rendering.UpdateFields_out_.pdf";
            doc.Save(dataDir);
            //ExEnd:UpdateDocFields
            Console.WriteLine("\nDocument fields updated successfully.\nFile saved at " + dataDir);
        }
    }
}
