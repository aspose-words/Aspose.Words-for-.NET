// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExRtfSaveOptions : ApiExampleBase
    {
        [TestCase(false)]
        [TestCase(true)]
        public void ExportImages(bool exportImagesForOldReaders)
        {
            //ExStart
            //ExFor:RtfSaveOptions
            //ExFor:RtfSaveOptions.ExportCompactSize
            //ExFor:RtfSaveOptions.ExportImagesForOldReaders
            //ExFor:RtfSaveOptions.SaveFormat
            //ExSummary:Shows how to save a document to .rtf with custom options.
            Document doc = new Document(MyDir + "Rendering.docx");

            // Create an "RtfSaveOptions" object to pass to the document's "Save" method to modify how we save it to an RTF.
            RtfSaveOptions options = new RtfSaveOptions();

            Assert.AreEqual(SaveFormat.Rtf, options.SaveFormat);

            // Set the "ExportCompactSize" property to "true" to
            // reduce the saved document's size at the cost of right-to-left text compatibility.
            options.ExportCompactSize = true;

            // Set the "ExportImagesFotOldReaders" property to "true" to use extra keywords to ensure that our document is
            // compatible with pre-Microsoft Word 97 readers and WordPad.
            // Set the "ExportImagesFotOldReaders" property to "false" to reduce the size of the document,
            // but prevent old readers from being able to read any non-metafile or BMP images that the document may contain.
            options.ExportImagesForOldReaders = exportImagesForOldReaders;

            doc.Save(ArtifactsDir + "RtfSaveOptions.ExportImages.rtf", options);
            //ExEnd

            if (exportImagesForOldReaders)
            {
                TestUtil.FileContainsString("nonshppict", ArtifactsDir + "RtfSaveOptions.ExportImages.rtf");
                TestUtil.FileContainsString("shprslt", ArtifactsDir + "RtfSaveOptions.ExportImages.rtf");
            }
            else
            {
                if (!IsRunningOnMono())
                {
                    Assert.Throws<AssertionException>(() =>
                        TestUtil.FileContainsString("nonshppict", ArtifactsDir + "RtfSaveOptions.ExportImages.rtf"));
                    Assert.Throws<AssertionException>(() =>
                        TestUtil.FileContainsString("shprslt", ArtifactsDir + "RtfSaveOptions.ExportImages.rtf"));
                }
            }
        }

        [TestCase(false), Category("SkipMono")]
        [TestCase(true), Category("SkipMono")]
        public void SaveImagesAsWmf(bool saveImagesAsWmf)
        {
            //ExStart
            //ExFor:RtfSaveOptions.SaveImagesAsWmf
            //ExSummary:Shows how to convert all images in a document to the Windows Metafile format as we save the document as an RTF.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Jpeg image:");
            Shape imageShape = builder.InsertImage(ImageDir + "Logo.jpg");

            Assert.AreEqual(ImageType.Jpeg, imageShape.ImageData.ImageType);

            builder.InsertParagraph();
            builder.Writeln("Png image:");
            imageShape = builder.InsertImage(ImageDir + "Transparent background logo.png");

            Assert.AreEqual(ImageType.Png, imageShape.ImageData.ImageType);

            // Create an "RtfSaveOptions" object to pass to the document's "Save" method to modify how we save it to an RTF.
            RtfSaveOptions rtfSaveOptions = new RtfSaveOptions();

            // Set the "SaveImagesAsWmf" property to "true" to convert all images in the document to WMF as we save it to RTF.
            // Doing so will help readers such as WordPad to read our document.
            // Set the "SaveImagesAsWmf" property to "false" to preserve the original format of all images in the document
            // as we save it to RTF. This will preserve the quality of the images at the cost of compatibility with older RTF readers.
            rtfSaveOptions.SaveImagesAsWmf = saveImagesAsWmf;

            doc.Save(ArtifactsDir + "RtfSaveOptions.SaveImagesAsWmf.rtf", rtfSaveOptions);

            doc = new Document(ArtifactsDir + "RtfSaveOptions.SaveImagesAsWmf.rtf");

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            if (saveImagesAsWmf)
            {
                Assert.AreEqual(ImageType.Wmf, ((Shape)shapes[0]).ImageData.ImageType);
                Assert.AreEqual(ImageType.Wmf, ((Shape)shapes[1]).ImageData.ImageType);
            }
            else
            {
                Assert.AreEqual(ImageType.Jpeg, ((Shape)shapes[0]).ImageData.ImageType);
                Assert.AreEqual(ImageType.Png, ((Shape)shapes[1]).ImageData.ImageType);
            }
            //ExEnd
        }
    }
}