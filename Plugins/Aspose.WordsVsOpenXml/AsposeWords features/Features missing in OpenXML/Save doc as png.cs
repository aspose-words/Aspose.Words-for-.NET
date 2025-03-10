// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features.Features_missing_in_OpenXML
{
    [TestFixture]
    public class SaveDocAsPng : TestUtil
    {
        [Test]
        public static void SaveDocAsPngFeature()
        {
            Document doc = new Document(MyDir + "Document.docx");

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);
            options.Resolution = 160;

            // Save each page of the document as Png.
            for (int i = 0; i < doc.PageCount; i++)
            {
                options.PageSet = new PageSet(i);
                doc.Save(ArtifactsDir + $"Save doc as png ({i}) - Aspose.Words.png", options);
            }
        }
    }
}
