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
using Aspose.Words.Fields;
using Aspose.Words.Layout;

namespace CSharp.Programming_With_Documents.Working_with_Fields
{
    class ConvertFieldsInDocument
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = _RunExamples.GetDataDir_WorkingWithFields();

            Document doc = new Document(dataDir + "TestFile.doc");

            // Pass the appropriate parameters to convert all IF fields encountered in the document (including headers and footers) to static text.
            FieldsHelper.ConvertFieldsToStaticText(doc, FieldType.FieldIf);

            // Save the document with fields transformed to disk.
            doc.Save(dataDir + "TestFileDocument Out.doc");

            Console.WriteLine("\nConverted fields to static text in the document successfully.\nFile saved at " + dataDir + "TestFileDocument Out.doc");
        }
    }
}
