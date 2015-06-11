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

namespace CSharp.Programming_With_Documents.Working_with_Tables
{
    class AutoFitTableToWindow
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = _RunExamples.GetDataDir_WorkingWithTables();

            // Open the document
            Document doc = new Document(dataDir + "TestFile.doc");

            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            // Autofit the first table to the page width.
            table.AutoFit(AutoFitBehavior.AutoFitToWindow);

            // Save the document to disk.
            doc.Save(dataDir + "TestFile.AutoFitToWindow Out.doc");

            Debug.Assert(doc.FirstSection.Body.Tables[0].PreferredWidth.Type == PreferredWidthType.Percent, "PreferredWidth type is not percent");
            Debug.Assert(doc.FirstSection.Body.Tables[0].PreferredWidth.Value == 100, "PreferredWidth value is different than 100");

            Console.WriteLine("\nAuto fit tables to window successfully.\nFile saved at " + dataDir + "TestFile.AutoFitToWindow Out.doc");
        }
    }
}
