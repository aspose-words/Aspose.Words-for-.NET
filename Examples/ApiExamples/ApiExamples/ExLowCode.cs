// Copyright (c) 2001-2024 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.LowCode;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    class ExLowCode : ApiExampleBase
    {
        [Test]
        public void MergeDocuments()
        {
            //ExStart
            //ExFor:Merger.Merge(String, String[])
            //ExFor:Merger.Merge(String[], MergeFormatMode)
            //ExFor:Merger.Merge(String, String[], SaveOptions, MergeFormatMode)
            //ExFor:Merger.Merge(String, String[], SaveFormat, MergeFormatMode)
            //ExFor:LowCode.MergeFormatMode
            //ExFor:LowCode.Merger
            //ExSummary:Shows how to merge documents into a single output document.
            //There is a several ways to merge documents:
            Merger.Merge(ArtifactsDir + "LowCode.MergeDocument.SimpleMerge.docx", new[] { MyDir + "Big document.docx", MyDir + "Tables.docx" });

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions { Password = "Aspose.Words" };
            Merger.Merge(ArtifactsDir + "LowCode.MergeDocument.SaveOptions.docx", new[] { MyDir + "Big document.docx", MyDir + "Tables.docx" }, saveOptions, MergeFormatMode.KeepSourceFormatting);

            Merger.Merge(ArtifactsDir + "LowCode.MergeDocument.SaveFormat.pdf", new[] { MyDir + "Big document.docx", MyDir + "Tables.docx" }, SaveFormat.Pdf, MergeFormatMode.KeepSourceLayout);

            Document doc = Merger.Merge(new[] { MyDir + "Big document.docx", MyDir + "Tables.docx" }, MergeFormatMode.MergeFormatting);
            doc.Save(ArtifactsDir + "LowCode.MergeDocument.DocumentInstance.docx");
            //ExEnd
        }

        [Test]
        public void MergeStreamDocument()
        {
            //ExStart
            //ExFor:Merger.Merge(Stream[], MergeFormatMode)
            //ExFor:Merger.Merge(Stream, Stream[], SaveOptions, MergeFormatMode)
            //ExFor:Merger.Merge(Stream, Stream[], SaveFormat)
            //ExSummary:Shows how to merge documents from stream into a single output document.
            //There is a several ways to merge documents from stream:
            using (FileStream firstStreamIn = new FileStream(MyDir + "Big document.docx", FileMode.Open, FileAccess.Read))
            {
                using (FileStream secondStreamIn = new FileStream(MyDir + "Tables.docx", FileMode.Open, FileAccess.Read))
                {
                    OoxmlSaveOptions saveOptions = new OoxmlSaveOptions { Password = "Aspose.Words" };
                    using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MergeStreamDocument.SaveOptions.docx", FileMode.Create, FileAccess.ReadWrite))
                        Merger.Merge(streamOut, new[] { firstStreamIn, secondStreamIn }, saveOptions, MergeFormatMode.KeepSourceFormatting);

                    using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MergeStreamDocument.SaveFormat.docx", FileMode.Create, FileAccess.ReadWrite))
                        Merger.Merge(streamOut, new[] { firstStreamIn, secondStreamIn }, SaveFormat.Docx);

                    Document doc = Merger.Merge(new[] { firstStreamIn, secondStreamIn }, MergeFormatMode.MergeFormatting);
                    doc.Save(ArtifactsDir + "LowCode.MergeStreamDocument.DocumentInstance.docx");
                }
            }
            //ExEnd
        }

        [Test]
        public void MergeDocumentInstances()
        {
            //ExStart:MergeDocumentInstances
            //GistId:e386727403c2341ce4018bca370a5b41
            //ExFor:Merger.Merge(Document[], MergeFormatMode)
            //ExSummary:Shows how to merge input documents to a single document instance.
            DocumentBuilder firstDoc = new DocumentBuilder();
            firstDoc.Font.Size = 16;
            firstDoc.Font.Color = Color.Blue;
            firstDoc.Write("Hello first word!");

            DocumentBuilder secondDoc = new DocumentBuilder();
            secondDoc.Write("Hello second word!");

            Document mergedDoc = Merger.Merge(new Document[] { firstDoc.Document, secondDoc.Document }, MergeFormatMode.KeepSourceLayout);
            Assert.AreEqual("Hello first word!\fHello second word!\f", mergedDoc.GetText());
            //ExEnd:MergeDocumentInstances
        }

        [Test]
        public void Convert()
        {
            //ExStart:Convert
            //GistId:708ce40a68fac5003d46f6b4acfd5ff1
            //ExFor:Converter.Convert(String, String)
            //ExFor:Converter.Convert(String, String, SaveFormat)
            //ExFor:Converter.Convert(String, String, SaveOptions)
            //ExSummary:Shows how to convert documents with a single line of code.
            Converter.Convert(MyDir + "Document.docx", ArtifactsDir + "LowCode.Convert.pdf");

            Converter.Convert(MyDir + "Document.docx", ArtifactsDir + "LowCode.Convert.rtf", SaveFormat.Rtf);

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions { Password = "Aspose.Words" };
            Converter.Convert(MyDir + "Document.doc", ArtifactsDir + "LowCode.Convert.docx", saveOptions);
            //ExEnd:Convert
        }

        [Test]
        public void ConvertStream()
        {
            //ExStart:ConvertStream
            //GistId:708ce40a68fac5003d46f6b4acfd5ff1
            //ExFor:Converter.Convert(Stream, Stream, SaveFormat)
            //ExFor:Converter.Convert(Stream, Stream, SaveOptions)
            //ExSummary:Shows how to convert documents with a single line of code (Stream).
            using (FileStream streamIn = new FileStream(MyDir + "Big document.docx", FileMode.Open, FileAccess.Read))
            {
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.ConvertStream.SaveFormat.docx", FileMode.Create, FileAccess.ReadWrite))
                    Converter.Convert(streamIn, streamOut, SaveFormat.Docx);

                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions { Password = "Aspose.Words" };
                using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.ConvertStream.SaveOptions.docx", FileMode.Create, FileAccess.ReadWrite))
                    Converter.Convert(streamIn, streamOut, saveOptions);
            }
            //ExEnd:ConvertStream
        }

        [Test]
        public void ConvertToImages()
        {
            //ExStart:ConvertToImages
            //GistId:708ce40a68fac5003d46f6b4acfd5ff1
            //ExFor:Converter.ConvertToImages(String, String)
            //ExFor:Converter.ConvertToImages(String, String, SaveFormat)
            //ExFor:Converter.ConvertToImages(String, String, SaveOptions)
            //ExSummary:Shows how to convert document to images.
            Converter.ConvertToImages(MyDir + "Big document.docx", ArtifactsDir + "LowCode.ConvertToImages.png");

            Converter.ConvertToImages(MyDir + "Big document.docx", ArtifactsDir + "LowCode.ConvertToImages.jpeg", SaveFormat.Jpeg);

            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.Png);
            imageSaveOptions.PageSet = new PageSet(1);
            Converter.ConvertToImages(MyDir + "Big document.docx", ArtifactsDir + "LowCode.ConvertToImages.png", imageSaveOptions);
            //ExEnd:ConvertToImages
        }

        [Test]
        public void ConvertToImagesStream()
        {
            //ExStart:ConvertToImagesStream
            //GistId:708ce40a68fac5003d46f6b4acfd5ff1
            //ExFor:Converter.ConvertToImages(String, SaveFormat)
            //ExFor:Converter.ConvertToImages(String, SaveOptions)
            //ExFor:Converter.ConvertToImages(Document, SaveFormat)
            //ExFor:Converter.ConvertToImages(Document, SaveOptions)
            //ExSummary:Shows how to convert document to images stream.
            Stream[] streams = Converter.ConvertToImages(MyDir + "Big document.docx", SaveFormat.Png);

            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.Png);
            imageSaveOptions.PageSet = new PageSet(1);
            streams = Converter.ConvertToImages(MyDir + "Big document.docx", imageSaveOptions);

            streams = Converter.ConvertToImages(new Document(MyDir + "Big document.docx"), SaveFormat.Png);

            streams = Converter.ConvertToImages(new Document(MyDir + "Big document.docx"), imageSaveOptions);
            //ExEnd:ConvertToImagesStream
        }

        [Test]
        public void ConvertToImagesFromStream()
        {
            //ExStart:ConvertToImagesFromStream
            //GistId:708ce40a68fac5003d46f6b4acfd5ff1
            //ExFor:Converter.ConvertToImages(Stream, SaveFormat)
            //ExFor:Converter.ConvertToImages(Stream, SaveOptions)
            //ExSummary:Shows how to convert document to images from stream.
            using (FileStream streamIn = new FileStream(MyDir + "Big document.docx", FileMode.Open, FileAccess.Read))
            {
                Stream[] streams = Converter.ConvertToImages(streamIn, SaveFormat.Jpeg);

                ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.Png);
                imageSaveOptions.PageSet = new PageSet(1);
                streams = Converter.ConvertToImages(streamIn, imageSaveOptions);
            }
            //ExEnd:ConvertToImagesFromStream
        }
    }
}
