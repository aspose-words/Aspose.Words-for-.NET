
using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using System.Diagnostics;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Tables
{
    class AutoFitTableToWindow
    {
        public static void Run()
        {
            //ExStart:AutoFitTableToPageWidth
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithTables();
            string fileName = "TestFile.doc";
            // Open the document
            Document doc = new Document(dataDir + fileName);

            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Autofit the first table to the page width.
            table.AutoFit(AutoFitBehavior.AutoFitToWindow);

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName); 
            // Save the document to disk.
            doc.Save(dataDir);

            Debug.Assert(doc.FirstSection.Body.Tables[0].PreferredWidth.Type == PreferredWidthType.Percent, "PreferredWidth type is not percent");
            Debug.Assert(doc.FirstSection.Body.Tables[0].PreferredWidth.Value == 100, "PreferredWidth value is different than 100");
            //ExEnd:AutoFitTableToPageWidth
            Console.WriteLine("\nAuto fit tables to window successfully.\nFile saved at " + dataDir);
        }
    }
}
