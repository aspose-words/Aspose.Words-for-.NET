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
using System.Collections;

namespace CSharp.Programming_With_Documents.Joining_and_Appending
{
    class PrependDocument
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = _RunExamples.GetDataDir_JoiningAndAppending();

            Document dstDoc = new Document(dataDir + "TestFile.Destination.doc");
            Document srcDoc = new Document(dataDir + "TestFile.Source.doc");

            // Append the source document to the destination document. This causes the result to have line spacing problems.
            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);

            // Instead prepend the content of the destination document to the start of the source document.
            // This results in the same joined document but with no line spacing issues.
            DoPrepend(srcDoc, dstDoc, ImportFormatMode.KeepSourceFormatting);

            // Save the document
            dstDoc.Save(dataDir + "TestFile.Prepend.doc");

            Console.WriteLine("\nDocument prepended successfully.\nFile saved at " + dataDir + "TestFile.Prepend.doc");
        }

        public static void DoPrepend(Document dstDoc, Document srcDoc, ImportFormatMode mode)
        {
            // Loop through all sections in the source document. 
            // Section nodes are immediate children of the Document node so we can just enumerate the Document.
            ArrayList sections = new ArrayList(srcDoc.Sections.ToArray());

            // Reverse the order of the sections so they are prepended to start of the destination document in the correct order.
            sections.Reverse();

            foreach (Section srcSection in sections)
            {
                // Import the nodes from the source document.
                Node dstSection = dstDoc.ImportNode(srcSection, true, mode);

                // Now the new section node can be prepended to the destination document.
                // Note how PrependChild is used instead of AppendChild. This is the only line changed compared 
                // to the original method.
                dstDoc.PrependChild(dstSection);
            }
        }
    }
}
