//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using System.Diagnostics;

namespace CSharp.Programming_Documents.Working_with_Tables
{
    class AutoFitTableToContents
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithTables();

            Document doc = new Document(dataDir + "TestFile.doc");

            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Auto fit the table to the cell contents
            table.AutoFit(AutoFitBehavior.AutoFitToContents);

            // Save the document to disk.
            doc.Save(dataDir + "TestFile.AutoFitToContents Out.doc");

            Debug.Assert(doc.FirstSection.Body.Tables[0].PreferredWidth.Type == PreferredWidthType.Auto, "PreferredWidth type is not auto");
            Debug.Assert(doc.FirstSection.Body.Tables[0].FirstRow.FirstCell.CellFormat.PreferredWidth.Type == PreferredWidthType.Auto, "PrefferedWidth on cell is not auto");
            Debug.Assert(doc.FirstSection.Body.Tables[0].FirstRow.FirstCell.CellFormat.PreferredWidth.Value == 0, "PreferredWidth value is not 0");

            Console.WriteLine("\nAuto fit tables to contents successfully.\nFile saved at " + dataDir + "TestFile.AutoFitToContents Out.doc");
        }
    }
}
