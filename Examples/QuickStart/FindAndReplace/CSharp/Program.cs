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

namespace FindAndReplaceExample
{
    public class Program
    {
        public static void Main()
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            // Open the document.
            Document doc = new Document(dataDir + "ReplaceSimple.doc");

            // Check the text of the document
            Console.WriteLine("Original document text: " + doc.Range.Text);

            // Replace the text in the document.
            doc.Range.Replace("_CustomerName_", "James Bond", false, false);

            // Check the replacement was made.
            Console.WriteLine("Document text after replace: " + doc.Range.Text);

            // Save the modified document.
            doc.Save(dataDir + "ReplaceSimple Out.doc");
        }
    }
}