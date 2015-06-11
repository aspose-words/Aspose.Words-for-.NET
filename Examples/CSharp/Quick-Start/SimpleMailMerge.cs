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
    class SimpleMailMerge
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = _RunExamples.GetDataDir_QuickStart();

            Document doc = new Document(dataDir + "MailMerge Template.doc");

            // Fill the fields in the document with user data.
            doc.MailMerge.Execute(
                new string[] { "FullName", "Company", "Address", "Address2", "City" },
                new object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });

            // Saves the document to disk.
            doc.Save(dataDir + "MailMerge Result Out.docx");

            Console.WriteLine("\nMail merge performed successfully.\nFile saved at " + dataDir + "MailMerge Result Out.docx");
        }
    }
}
