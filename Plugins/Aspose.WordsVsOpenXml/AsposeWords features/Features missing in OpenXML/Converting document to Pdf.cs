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
    public class ConvertingDocumentToPdf : TestUtil
    {
        [Test]
        public static void ConvertingDocumentToPdfFeature()
        {
            Document doc = new Document(MyDir + "Document.docx");
            doc.Save(ArtifactsDir + "Converting document to Pdf - Aspose.Words.pdf");
        }
    }
}
