//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Words;
using System;

namespace CSharp.Quick_Start
{
    class LoadAndSaveToDisk
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_QuickStart();

            // Load the document from the absolute path on disk.
            Document doc = new Document(dataDir + "Document.doc");

            // Save the document as DOCX document.");
            doc.Save(dataDir + "Document Out.docx");

            Console.WriteLine("\nExisting document loaded and saved successfully.\nFile saved at " + dataDir + "HelloWorld Out.docx");
        }
    }
}
