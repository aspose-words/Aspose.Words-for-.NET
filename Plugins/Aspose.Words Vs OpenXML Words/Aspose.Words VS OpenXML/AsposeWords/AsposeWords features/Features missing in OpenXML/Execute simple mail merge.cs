// Copyright (c) Aspose 2002-2021. All Rights Reserved.

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
                new string[] { "Name", "City" },
                new object[] { "Zeeshan", "Islamabad" });

            // Send the document in Word format to the client browser with an option to save to disk or open inside the current browser.
            doc.Save(ArtifactsDir + "Execute simple mail merge - Aspose.Words.docx");
        }
    }
}
