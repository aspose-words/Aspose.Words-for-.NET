// Copyright (c) Aspose 2002-2021. All Rights Reserved.

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
                doc.Save(string.Format(ArtifactsDir + i + " Save doc as png - Aspose.Words.png", i), options);
            }
        }
    }
}
