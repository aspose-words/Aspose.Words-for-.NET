// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features.Features_missing_in_OpenXML
{
    [TestFixture]
    public class ExecuteSimpleMailMerge : TestUtil
    {
        [Test]
        public static void ExecuteSimpleMailMergeFeature()
        {
            Document doc = new Document(MyDir + "Mail merge.docx");

            // Fill the fields in the document with user data.
            doc.MailMerge.Execute(
                ["Name", "City"],
                ["Zeeshan", "Islamabad"]);

            // Send the document in Word format to the client browser with an option to save to disk or open inside the current browser.
            doc.Save(ArtifactsDir + "Execute simple mail merge - Aspose.Words.docx");
        }
    }
}
