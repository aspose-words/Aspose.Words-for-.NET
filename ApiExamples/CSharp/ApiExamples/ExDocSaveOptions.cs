// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExDocSaveOptions : ApiExampleBase
    {
        [Test]
        public void SaveAsDoc()
        {
            //ExStart
            //ExFor:DocSaveOptions
            //ExFor:DocSaveOptions.#ctor
            //ExFor:DocSaveOptions.#ctor(SaveFormat)
            //ExFor:DocSaveOptions.Password
            //ExFor:DocSaveOptions.SaveFormat
            //ExFor:DocSaveOptions.SaveRoutingSlip
            //ExSummary:Shows how to set save options for classic Microsoft Word document versions.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Hello world!");

            // DocSaveOptions only applies to Doc and Dot save formats
            DocSaveOptions options = new DocSaveOptions(SaveFormat.Doc);

            // Set a password with which the document will be encrypted, and which will be required to open it
            options.Password = "MyPassword";

            // If the document contains a routing slip, we can preserve it while saving by setting this flag to true
            options.SaveRoutingSlip = true;

            doc.Save(ArtifactsDir + "DocSaveOptions.SaveAsDoc.doc", options);
            //ExEnd

            Assert.Throws<IncorrectPasswordException>(() => doc = new Document(ArtifactsDir + "DocSaveOptions.SaveAsDoc.doc"));

            LoadOptions loadOptions = new LoadOptions("MyPassword");
            doc = new Document(ArtifactsDir + "DocSaveOptions.SaveAsDoc.doc", loadOptions);

            Assert.AreEqual("Hello world!", doc.GetText().Trim());
        }

        [Test]
        public void TempFolder()
        {
            //ExStart
            //ExFor:SaveOptions.TempFolder
            //ExSummary:Shows how to save a document using temporary files.
            Document doc = new Document(MyDir + "Rendering.docx");

            // We can use a SaveOptions object to set the saving method of a document from a MemoryStream to temporary files
            // While saving, the files will briefly pop up in the folder we set as the TempFolder attribute below
            // Doing this will free up space in the memory that the stream would usually occupy
            DocSaveOptions options = new DocSaveOptions();
            options.TempFolder = ArtifactsDir + "TempFiles";

            // Ensure that the directory exists and save
            Directory.CreateDirectory(options.TempFolder);

            doc.Save(ArtifactsDir + "DocSaveOptions.TempFolder.doc", options);
            //ExEnd
        }

        [Test]
        public void PictureBullets()
        {
            //ExStart
            //ExFor:DocSaveOptions.SavePictureBullet
            //ExSummary:Shows how to remove PictureBullet data from the document.
            Document doc = new Document(MyDir + "Image bullet points.docx");
            Assert.NotNull(doc.Lists[0].ListLevels[0].ImageData); //ExSkip

            // Word 97 cannot work correctly with PictureBullet data
            // To remove PictureBullet data, set the option to "false"
            DocSaveOptions saveOptions = new DocSaveOptions(SaveFormat.Doc);
            saveOptions.SavePictureBullet = false;

            doc.Save(ArtifactsDir + "DocSaveOptions.PictureBullets.doc", saveOptions);
            //ExEnd

            doc = new Document(ArtifactsDir + "DocSaveOptions.PictureBullets.doc");

            Assert.Null(doc.Lists[0].ListLevels[0].ImageData);
        }

        [TestCase(true)]
        [TestCase(false)]
        public void UpdateLastPrintedProperty(bool isUpdateLastPrintedProperty)
        {
            //ExStart
            //ExFor:SaveOptions.UpdateLastPrintedProperty
            //ExSummary:Shows how to update BuiltInDocumentProperties.LastPrinted property before saving.
            Document doc = new Document();

            // Aspose.Words update BuiltInDocumentProperties.LastPrinted property by default
            DocSaveOptions saveOptions = new DocSaveOptions();
            saveOptions.UpdateLastPrintedProperty = isUpdateLastPrintedProperty;

            doc.Save(ArtifactsDir + "DocSaveOptions.UpdateLastPrintedProperty.docx", saveOptions);
            //ExEnd

            doc = new Document(ArtifactsDir + "DocSaveOptions.UpdateLastPrintedProperty.docx");

            Assert.AreNotEqual(isUpdateLastPrintedProperty, DateTime.Parse("1/1/0001 00:00:00") == doc.BuiltInDocumentProperties.LastPrinted.Date);
        }
    }
}