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
using Aspose.Words.MailMerging;
using Aspose.Words.Saving;
using System.Text;

namespace CSharp.Loading_Saving
{
    class LoadTxt
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            // The encoding of the text file is automatically detected.
            Document doc = new Document(dataDir + "LoadTxt.txt");

            // Save as any Aspose.Words supported format, such as DOCX.
            doc.Save(dataDir + "LoadTxt Out.docx");

            Console.WriteLine("\nText document loaded successfully.\nFile saved at " + dataDir + "LoadTxt Out.docx");
        }
    }
}
