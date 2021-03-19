// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features.Features_missing_in_OpenXML
{
    [TestFixture]
    public class ConvertingDocumentToPdf : TestUtil
    {
        [Test]
        public static void ConvertingDocumentToPdfFeature()
        {
            Document doc = new Document(MyDir + "Document.docx");

            doc.Save(ArtifactsDir + "Converting document to Pdf - Aspose.Words.pdf", SaveFormat.Pdf);
        }
    }
}
