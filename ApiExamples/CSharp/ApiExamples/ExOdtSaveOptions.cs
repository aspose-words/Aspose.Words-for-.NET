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
    internal class ExOdtSaveOptions : ApiExampleBase
    {
        [Test]
        public void MeasureUnit()
        {
            //ExStart
            //ExFor:OdtSaveOptions
            //ExFor:OdtSaveOptions.#ctor
            //ExFor:OdtSaveOptions.IsStrictSchema11
            //ExFor:OdtSaveOptions.MeasureUnit
            //ExFor:OdtSaveMeasureUnit
            //ExSummary:Shows how to work with units of measure of document content.
            Document doc = new Document(MyDir + "Rendering.doc");

            // Open Office uses centimeters, MS Office uses inches
            OdtSaveOptions saveOptions = new OdtSaveOptions
            {
                MeasureUnit = OdtSaveMeasureUnit.Inches,
                IsStrictSchema11 = true
            };

            doc.Save(ArtifactsDir + "OdtSaveOptions.MeasureUnit.odt", saveOptions);
            //ExEnd
        }

        [Test]
        [TestCase(SaveFormat.Odt)]
        [TestCase(SaveFormat.Ott)]
        public void Encrypt(SaveFormat saveFormat)
        {
            //ExStart
            //ExFor:OdtSaveOptions.#ctor(SaveFormat)
            //ExFor:OdtSaveOptions.Password
            //ExFor:OdtSaveOptions.SaveFormat
            //ExSummary:Shows how to encrypted your odt/ott documents with a password.
            Document doc = new Document(MyDir + "Document.docx");

            OdtSaveOptions saveOptions = new OdtSaveOptions(saveFormat);
            saveOptions.Password = "@sposeEncrypted_1145";

            // Saving document using password property of OdtSaveOptions
            doc.Save(ArtifactsDir + "OdtSaveOptions.Encrypt" +
                     FileFormatUtil.SaveFormatToExtension(saveFormat), saveOptions);
            //ExEnd

            // Check that all documents are encrypted with a password
            FileFormatInfo docInfo = FileFormatUtil.DetectFileFormat(
                ArtifactsDir + "OdtSaveOptions.Encrypt" +
                FileFormatUtil.SaveFormatToExtension(saveFormat));
            Assert.IsTrue(docInfo.IsEncrypted);
        }

        [Test]
        [TestCase(SaveFormat.Odt)]
        [TestCase(SaveFormat.Ott)]
        public void WorkWithEncryptedDocument(SaveFormat saveFormat)
        {
            //ExStart
            //ExFor:OdtSaveOptions.#ctor(String)
            //ExSummary:Shows how to load and change odt/ott encrypted document.
            Document doc = new Document(MyDir + "Encrypted" +
                                        FileFormatUtil.SaveFormatToExtension(saveFormat),
                new LoadOptions("@sposeEncrypted_1145"));

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();
            builder.Writeln("Encrypted document after changes.");

            // Saving document using new instance of OdtSaveOptions
            doc.Save(ArtifactsDir + "OdtSaveOptions.WorkWithEncryptedDocument" +
                     FileFormatUtil.SaveFormatToExtension(saveFormat), new OdtSaveOptions("@sposeEncrypted_1145"));
            //ExEnd

            // Check that document is still encrypted with a password
            FileFormatInfo docInfo = FileFormatUtil.DetectFileFormat(
                ArtifactsDir + "OdtSaveOptions.WorkWithEncryptedDocument" +
                FileFormatUtil.SaveFormatToExtension(saveFormat));
            Assert.IsTrue(docInfo.IsEncrypted);
        }
    }
}