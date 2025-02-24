// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features.Features_missing_in_OpenXML
{
    [TestFixture]
    public class PrintDocumentViaXpsApi : TestUtil
    {
        [Test, Ignore("Run only when a printer driver installed")]
        public static void PrintDocumentViaXpsApiFeature()
        {
            try
            {
                Document document = new Document(MyDir + "Document.docx");

                // Specify the name of the printer you want to print to.
                const string printerName = @"\\COMPANY\Zeeshan MFC-885CW Printer";
                // Print the document.
                XpsPrintHelper.Print(document, printerName, "test", true);
                
                Console.WriteLine("Printed successfully.");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
        }
    }
}
