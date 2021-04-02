// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class ExtractImageFromWordDocument : TestUtil
    {
        [Test]
        public void ExtractImageFromWordDocumentFeature()
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(MyDir + "Extract image.docx", false))
            {
                int imgCount = doc.MainDocumentPart.GetPartsOfType<ImagePart>().Count();

                if (imgCount > 0)
                {
                    List<ImagePart> imgParts = new List<ImagePart>(doc.MainDocumentPart.ImageParts);

                    foreach (ImagePart imgPart in imgParts)
                    {
                        Image img = Image.FromStream(imgPart.GetStream());
                        string imgfileName = imgPart.Uri.OriginalString.Substring(imgPart.Uri.OriginalString.LastIndexOf("/") + 1);

                        img.Save(ArtifactsDir + imgfileName);
                    }
                }
            }
        }
    }
}
