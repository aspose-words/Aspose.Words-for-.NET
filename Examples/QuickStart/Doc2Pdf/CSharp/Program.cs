//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Words;

namespace Doc2PdfExample
{
    public class Program
    {
        public static void Main()
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            // Load the document from disk.
            Document doc = new Document(dataDir + "Template.doc");

            // Save the document in PDF format.
            doc.Save(dataDir + "Doc2PdfSave Out.pdf");
        }
    }
}