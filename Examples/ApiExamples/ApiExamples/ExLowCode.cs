// Copyright (c) 2001-2023 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
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
            //ExSummary:Shows how to merge documents into a single output document.
            //There is a several ways to merge documents:
            Merger.Merge(ArtifactsDir + "LowCode.MergeDocument.SimpleMerge.docx", new[] { MyDir + "Big document.docx", MyDir + "Tables.docx" });

            Merger.Merge(ArtifactsDir + "LowCode.MergeDocument.SaveOptions.docx", new[] { MyDir + "Big document.docx", MyDir + "Tables.docx" }, new OoxmlSaveOptions() { Password = "Aspose.Words" }, MergeFormatMode.KeepSourceFormatting);

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
                    using (FileStream streamOut = new FileStream(ArtifactsDir + "LowCode.MergeStreamDocument.SaveOptions.docx", FileMode.Create, FileAccess.ReadWrite))
                        Merger.Merge(streamOut, new[] { firstStreamIn, secondStreamIn }, new OoxmlSaveOptions() { Password = "Aspose.Words" }, MergeFormatMode.KeepSourceFormatting);

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
            //ReleaseVersion:23.12
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
    }
}
