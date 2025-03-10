// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.IO;
using DocumentFormat.OpenXml.Packaging;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class ExtractImages : TestUtil
    {
        [Test]
        //ExStart:ExtractDocumentImagesOpenXml
        //GistId:49853740887e9a787707ab66bb4ec5e2
        public void ExtractDocumentImagesOpenXml()
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(MyDir + "Extract image.docx", false))
            {
                MainDocumentPart mainPart = wordDocument.MainDocumentPart;

                // Loop through each ImagePart in the document.
                int imageIndex = 0;
                foreach (var imagePart in mainPart.ImageParts)
                {
                    // Get the image extension.
                    string imageExtension = GetImageExtension(imagePart.ContentType);
                    if (imageExtension != null)
                    {
                        // Create a file name for the image.
                        string imageFileName = Path.Combine(ArtifactsDir, $"ExtractDocumentImages.{imageIndex}.OpenXML{imageExtension}");

                        // Save the image to the output directory.
                        using (var stream = imagePart.GetStream())
                        using (var fileStream = new FileStream(imageFileName, FileMode.Create, FileAccess.Write))
                            stream.CopyTo(fileStream);

                        imageIndex++;
                    }
                }
            }
        }

        static string GetImageExtension(string contentType)
        {
            switch (contentType)
            {
                case "image/jpeg":
                    return ".jpg";
                case "image/png":
                    return ".png";
                case "image/gif":
                    return ".gif";
                case "image/bmp":
                    return ".bmp";
                default:
                    return null; // Unsupported image type.
            }
        }
        //ExEnd:ExtractDocumentImagesOpenXml
    }
}
