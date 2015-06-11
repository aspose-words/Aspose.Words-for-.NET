//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;

using Aspose.Words;

namespace CSharp.Programming_With_Documents.Joining_and_Appending
{
    class LinkHeadersFooters
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = _RunExamples.GetDataDir_JoiningAndAppending();

            Document dstDoc = new Document(dataDir + "TestFile.Destination.doc");
            Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

            // Set the appended document to appear on a new page.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.NewPage;

            // Link the headers and footers in the source document to the previous section. 
            // This will override any headers or footers already found in the source document. 
            srcDoc.FirstSection.HeadersFooters.LinkToPrevious(true);

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            dstDoc.Save(dataDir + "TestFile.LinkHeadersFooters Out.doc");

            Console.WriteLine("\nDocument appended successfully with linked header footers.\nFile saved at " + dataDir + "TestFile.LinkHeadersFooters Out.doc");
        }
    }
}
