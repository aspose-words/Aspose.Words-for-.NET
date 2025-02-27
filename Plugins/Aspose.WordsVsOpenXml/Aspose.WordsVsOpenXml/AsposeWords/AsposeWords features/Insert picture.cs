// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Drawing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class InsertPicture: TestUtil
    {
        [Test]
        public void InsertImageAsposeWords()
        {
            //ExStart:InsertImageAsposeWords
            //GistDesc:Insert image using C#
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertImage(MyDir + "Aspose.Words.png",
                RelativeHorizontalPosition.Margin,
                100,
                RelativeVerticalPosition.Margin,
                100,
                200,
                100,
                WrapType.Square);

            doc.Save(ArtifactsDir + "Insert image - Aspose.Words.docx");
            //ExEnd:InsertImageAsposeWords
        }
    }
}
