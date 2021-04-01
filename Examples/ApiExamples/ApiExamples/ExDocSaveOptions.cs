// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Loading;
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
            //ExSummary:Shows how to set save options for older Microsoft Word formats.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Hello world!");

            DocSaveOptions options = new DocSaveOptions(SaveFormat.Doc);
            
            // Set a password which will protect the loading of the document by Microsoft Word or Aspose.Words.
            // Note that this does not encrypt the contents of the document in any way.
            options.Password = "MyPassword";

            // If the document contains a routing slip, we can preserve it while saving by setting this flag to true.
            options.SaveRoutingSlip = true;

            doc.Save(ArtifactsDir + "DocSaveOptions.SaveAsDoc.doc", options);

            // To be able to load the document,
            // we will need to apply the password we specified in the DocSaveOptions object in a LoadOptions object.
            Assert.Throws<IncorrectPasswordException>(() => doc = new Document(ArtifactsDir + "DocSaveOptions.SaveAsDoc.doc"));

            LoadOptions loadOptions = new LoadOptions("MyPassword");
            doc = new Document(ArtifactsDir + "DocSaveOptions.SaveAsDoc.doc", loadOptions);

            Assert.AreEqual("Hello world!", doc.GetText().Trim());
            //ExEnd
        }

        [Test]
        public void TempFolder()
        {
            //ExStart
            //ExFor:SaveOptions.TempFolder
            //ExSummary:Shows how to use the hard drive instead of memory when saving a document.
            Document doc = new Document(MyDir + "Rendering.docx");

            // When we save a document, various elements are temporarily stored in memory as the save operation is taking place.
            // We can use this option to use a temporary folder in the local file system instead,
            // which will reduce our application's memory overhead.
            DocSaveOptions options = new DocSaveOptions();
            options.TempFolder = ArtifactsDir + "TempFiles";

            // The specified temporary folder must exist in the local file system before the save operation.
            Directory.CreateDirectory(options.TempFolder);

            doc.Save(ArtifactsDir + "DocSaveOptions.TempFolder.doc", options);

            // The folder will persist with no residual contents from the load operation.
            Assert.That(Directory.GetFiles(options.TempFolder), Is.Empty);
            //ExEnd
        }

        [Test]
        public void PictureBullets()
        {
            //ExStart
            //ExFor:DocSaveOptions.SavePictureBullet
            //ExSummary:Shows how to omit PictureBullet data from the document when saving.
            Document doc = new Document(MyDir + "Image bullet points.docx");
            Assert.NotNull(doc.Lists[0].ListLevels[0].ImageData); //ExSkip

            // Some word processors, such as Microsoft Word 97, are incompatible with PictureBullet data.
            // By setting a flag in the SaveOptions object,
            // we can convert all image bullet points to ordinary bullet points while saving.
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
            //ExSummary:Shows how to update a document's "Last printed" property when saving.
            Document doc = new Document();
            doc.BuiltInDocumentProperties.LastPrinted = new DateTime(2019, 12, 20);

            // This flag determines whether the last printed date, which is a built-in property, is updated.
            // If so, then the date of the document's most recent save operation
            // with this SaveOptions object passed as a parameter is used as the print date.
            DocSaveOptions saveOptions = new DocSaveOptions();
            saveOptions.UpdateLastPrintedProperty = isUpdateLastPrintedProperty;

            // In Microsoft Word 2003, this property can be found via File -> Properties -> Statistics -> Printed.
            // It can also be displayed in the document's body by using a PRINTDATE field.
            doc.Save(ArtifactsDir + "DocSaveOptions.UpdateLastPrintedProperty.doc", saveOptions);

            // Open the saved document, then verify the value of the property.
            doc = new Document(ArtifactsDir + "DocSaveOptions.UpdateLastPrintedProperty.doc");

            Assert.AreNotEqual(isUpdateLastPrintedProperty, new DateTime(2019, 12, 20) == doc.BuiltInDocumentProperties.LastPrinted);
            //ExEnd
        }

        [TestCase(true)]
        [TestCase(false)]
        public void UpdateCreatedTimeProperty(bool isUpdateCreatedTimeProperty)
        {
            //ExStart
            //ExFor:SaveOptions.UpdateLastPrintedProperty
            //ExSummary:Shows how to update a document's "CreatedTime" property when saving.
            Document doc = new Document();
            doc.BuiltInDocumentProperties.CreatedTime = new DateTime(2019, 12, 20);

            // This flag determines whether the created time, which is a built-in property, is updated.
            // If so, then the date of the document's most recent save operation
            // with this SaveOptions object passed as a parameter is used as the created time.
            DocSaveOptions saveOptions = new DocSaveOptions();
            saveOptions.UpdateCreatedTimeProperty = isUpdateCreatedTimeProperty;

            doc.Save(ArtifactsDir + "DocSaveOptions.UpdateCreatedTimeProperty.docx", saveOptions);

            // Open the saved document, then verify the value of the property.
            doc = new Document(ArtifactsDir + "DocSaveOptions.UpdateCreatedTimeProperty.docx");

            Assert.AreNotEqual(isUpdateCreatedTimeProperty, new DateTime(2019, 12, 20) == doc.BuiltInDocumentProperties.CreatedTime);
            //ExEnd
        }

        [TestCase(false)]
        [TestCase(true)]
        public void AlwaysCompressMetafiles(bool compressAllMetafiles)
        {
            //ExStart
            //ExFor:DocSaveOptions.AlwaysCompressMetafiles
            //ExSummary:Shows how to change metafiles compression in a document while saving.
            // Open a document that contains a Microsoft Equation 3.0 formula.
            Document doc = new Document(MyDir + "Microsoft equation object.docx");

            // When we save a document, smaller metafiles are not compressed for performance reasons.
            // We can set a flag in a SaveOptions object to compress every metafile when saving.
            // Some editors such as LibreOffice cannot read uncompressed metafiles.
            DocSaveOptions saveOptions = new DocSaveOptions();
            saveOptions.AlwaysCompressMetafiles = compressAllMetafiles;

            doc.Save(ArtifactsDir + "DocSaveOptions.AlwaysCompressMetafiles.docx", saveOptions);

            if (compressAllMetafiles)
                Assert.That(10000, Is.LessThan(new FileInfo(ArtifactsDir + "DocSaveOptions.AlwaysCompressMetafiles.docx").Length));
            else
                Assert.That(30000, Is.AtLeast(new FileInfo(ArtifactsDir + "DocSaveOptions.AlwaysCompressMetafiles.docx").Length));
            //ExEnd
        }
    }
}