
using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using System.Diagnostics;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Tables
{
    class AutoFitTableToFixedColumnWidths
    {
        public static void Run()
        {
            //ExStart:AutoFitTableToFixedColumnWidths
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithTables();
            string fileName = "TestFile.doc";
            Document doc = new Document(dataDir + fileName);

            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Disable autofitting on this table.
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            // Save the document to disk.
            doc.Save(dataDir);
            //ExEnd

            Debug.Assert(doc.FirstSection.Body.Tables[0].PreferredWidth.Type == PreferredWidthType.Auto, "PreferredWidth type is not auto");
            Debug.Assert(doc.FirstSection.Body.Tables[0].PreferredWidth.Value == 0, "PreferredWidth value is not 0");
            Debug.Assert(doc.FirstSection.Body.Tables[0].FirstRow.FirstCell.CellFormat.Width == 69.2, "Cell width is not correct.");
            //ExEnd:AutoFitTableToFixedColumnWidths
            Console.WriteLine("\nAuto fit tables to fixed column widths successfully.\nFile saved at " + dataDir);
        }
    }
}
