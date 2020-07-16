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
            //ExSummary:Shows how to set save options for older Microsoft Word formats.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Hello world!");

            DocSaveOptions options = new DocSaveOptions(SaveFormat.Doc);

            // Set a password with which the document will be protected during loading by Microsoft Word or Aspose.Words.
            // Note that the document is not in any way encrypted.
            options.Password = "MyPassword";

            // If the document contains a routing slip, we can preserve it while saving by setting this flag to true.
            options.SaveRoutingSlip = true;

            doc.Save(ArtifactsDir + "DocSaveOptions.SaveAsDoc.doc", options);

            // In order to be able to load the document,
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

            // This flag determines whether the last printed date, which is stored in the document's built-in properties, is updated.
            // If it is, then the date when the document was saved with this SaveOptions object is used as the print date.
            DocSaveOptions saveOptions = new DocSaveOptions();
            saveOptions.UpdateLastPrintedProperty = isUpdateLastPrintedProperty;

            // In Microsoft Word 2003, this property can be found via File -> Properties -> Statistics -> Printed.
            // It can also be displayed in the document's body by using a PRINTDATE field.
            doc.Save(ArtifactsDir + "DocSaveOptions.UpdateLastPrintedProperty.doc", saveOptions);

            // Open the saved document, then verify the value of the property.
            doc = new Document(ArtifactsDir + "DocSaveOptions.UpdateLastPrintedProperty.doc");

            Assert.AreNotEqual(isUpdateLastPrintedProperty, (DateTime.MinValue.Date == doc.BuiltInDocumentProperties.LastPrinted));
            //ExEnd
        }
    }
}