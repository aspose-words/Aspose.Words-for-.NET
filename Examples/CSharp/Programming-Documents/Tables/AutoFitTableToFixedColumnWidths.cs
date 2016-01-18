
using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using System.Diagnostics;

namespace CSharp.Programming_Documents.Working_with_Tables
{
    class AutoFitTableToFixedColumnWidths
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithTables();

            Document doc = new Document(dataDir + "TestFile.doc");

            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Disable autofitting on this table.
            table.AutoFit(AutoFitBehavior.FixedColumnWidths);

            // Save the document to disk.
            doc.Save(dataDir + "TestFile.FixedWidth Out.doc");
            //ExEnd

            Debug.Assert(doc.FirstSection.Body.Tables[0].PreferredWidth.Type == PreferredWidthType.Auto, "PreferredWidth type is not auto");
            Debug.Assert(doc.FirstSection.Body.Tables[0].PreferredWidth.Value == 0, "PreferredWidth value is not 0");
            Debug.Assert(doc.FirstSection.Body.Tables[0].FirstRow.FirstCell.CellFormat.Width == 69.2, "Cell width is not correct.");

            Console.WriteLine("\nAuto fit tables to fixed column widths successfully.\nFile saved at " + dataDir + "TestFile.FixedWidth Out.doc");
        }
    }
}
