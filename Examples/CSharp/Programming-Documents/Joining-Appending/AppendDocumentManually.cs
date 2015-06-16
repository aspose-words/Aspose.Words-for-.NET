﻿//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;

using Aspose.Words;

namespace CSharp.Programming_Documents.Joining_and_Appending
{
    class AppendDocumentManually
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_JoiningAndAppending();

            Document dstDoc = new Document(dataDir + "TestFile.Destination.doc");
            Document srcDoc = new Document(dataDir + "TestFile.Source.doc");
            ImportFormatMode mode = ImportFormatMode.KeepSourceFormatting;

            // Loop through all sections in the source document. 
            // Section nodes are immediate children of the Document node so we can just enumerate the Document.
            foreach (Section srcSection in srcDoc)
            {
                // Because we are copying a section from one document to another, 
                // it is required to import the Section node into the destination document.
                // This adjusts any document-specific references to styles, lists, etc.
                //
                // Importing a node creates a copy of the original node, but the copy
                // is ready to be inserted into the destination document.
                Node dstSection = dstDoc.ImportNode(srcSection, true, mode);

                // Now the new section node can be appended to the destination document.
                dstDoc.AppendChild(dstSection);
            }

            // Save the joined document
            dstDoc.Save(dataDir + "TestFile.Append Manual Out.doc");

            Console.WriteLine("\nDocument appended successfully with manual append operation.\nFile saved at " + dataDir + "TestFile.Append Manual Out.pdf");
        }
    }
}
