using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Styles
{
    class CopyStyles
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithStyles();

            // ExStart:CopyStylesFromDocument
            string fileName = dataDir + "template.docx";
            Document doc = new Document(fileName);

            // Open the document.
            Document target = new Document(dataDir + "TestFile.doc");
            target.CopyStylesFromTemplate(doc);
            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            doc.Save(dataDir);
            // ExEnd:CopyStylesFromDocument 
            Console.WriteLine("\nStyles are copied from document successfully.\nFile saved at " + dataDir);
        }
    }
}
