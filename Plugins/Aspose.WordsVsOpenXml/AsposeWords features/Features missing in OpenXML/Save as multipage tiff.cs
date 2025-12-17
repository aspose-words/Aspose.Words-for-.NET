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
    public class SaveAsMultiPageTiff : TestUtil
    {
        [Test]
        public static void SaveAsMultiPageTiffFeature()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Tiff);
            options.PageSet = new PageSet(new PageRange(0, doc.PageCount));
            options.TiffCompression = TiffCompression.Ccitt4;
            options.Resolution = 160;

            doc.Save(ArtifactsDir + "Save as multipage tiff - Aspose.Words.tiff", options);
        }
    }
}
