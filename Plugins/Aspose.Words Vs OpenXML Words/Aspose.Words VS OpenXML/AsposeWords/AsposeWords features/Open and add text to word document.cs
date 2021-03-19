// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System.Drawing;
using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class OpenAndAddTextToWordDocument : TestUtil
    {
        [Test]
        public void OpenAndAddTextToWordDocumentFeature()
        {
            Document doc = new Document(MyDir + "Document.docx");
            DocumentBuilder builder = new DocumentBuilder(doc);

            global::Aspose.Words.Font font = builder.Font;
            font.Size = 16;
            font.Bold = true;
            font.Color = Color.Blue;
            font.Name = "Arial";
            font.Underline = Underline.Dash;

            builder.Write("Formatted text.");

            doc.Save(ArtifactsDir + "Open and add text - Aspose.Words.docx");
        }
    }
}
