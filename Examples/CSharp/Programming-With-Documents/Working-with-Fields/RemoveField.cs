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
    class RemoveField
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = _RunExamples.GetDataDir_WorkingWithFields();

            Document doc = new Document(dataDir + "Field.RemoveField.doc");

            Field field = doc.Range.Fields[0];
            // Calling this method completely removes the field from the document.
            field.Remove();

            Console.WriteLine("\nRemoved field from the document successfully.");
        }
    }
}
