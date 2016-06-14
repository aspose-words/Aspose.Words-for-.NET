
using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using System.Diagnostics;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Tables
{
    class AutoFitTableToContents
    {
        public static void Run()
        {
            //ExStart:AutoFitTableToContents
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithTables();

            string fileName = "TestFile.doc";
            Document doc = new Document(dataDir + fileName);

            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Auto fit the table to the cell contents
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            // Save the document to disk.
            doc.Save(dataDir);

            Debug.Assert(doc.FirstSection.Body.Tables[0].PreferredWidth.Type == PreferredWidthType.Auto, "PreferredWidth type is not auto");
            Debug.Assert(doc.FirstSection.Body.Tables[0].FirstRow.FirstCell.CellFormat.PreferredWidth.Type == PreferredWidthType.Auto, "PrefferedWidth on cell is not auto");
            Debug.Assert(doc.FirstSection.Body.Tables[0].FirstRow.FirstCell.CellFormat.PreferredWidth.Value == 0, "PreferredWidth value is not 0");
            //ExEnd:AutoFitTableToContents
            Console.WriteLine("\nAuto fit tables to contents successfully.\nFile saved at " + dataDir);
        }
    }
}
