// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExRtfSaveOptions : ApiExampleBase
    {
        [Test]
        public void ExportImages()
        {
            //ExStart
            //ExFor:RtfSaveOptions
            //ExFor:RtfSaveOptions.ExportCompactSize
            //ExFor:RtfSaveOptions.ExportImagesForOldReaders
            //ExFor:RtfSaveOptions.SaveFormat
            //ExSummary:Shows how to save a document to .rtf with custom options.
            // Open a document with images
            Document doc = new Document(MyDir + "Rendering.doc");

            // Configure a RtfSaveOptions instance to make our output document more suitable for older devices
            RtfSaveOptions options = new RtfSaveOptions
            {
                SaveFormat = SaveFormat.Rtf,
                ExportCompactSize = true,
                ExportImagesForOldReaders = true
            };

            doc.Save(ArtifactsDir + "RtfSaveOptions.ExportImages.rtf", options);
            //ExEnd
        }
    }
}