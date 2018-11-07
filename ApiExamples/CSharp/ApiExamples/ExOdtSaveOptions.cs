// Copyright (c) 2001-2018 Aspose Pty Ltd. All Rights Reserved.
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
        public void MeasureUnitOption()
        {
            //ExStart
            //ExFor:OdtSaveOptions.MeasureUnit
            //ExFor:OdtSaveMeasureUnit
            //ExSummary:Shows how to work with units of measure of document content
            Document doc = new Document(MyDir + "OdtSaveOptions.MeasureUnit.docx");

            // Open Office uses centimeters, MS Office uses inches
            OdtSaveOptions saveOptions = new OdtSaveOptions();
            saveOptions.MeasureUnit = OdtSaveMeasureUnit.Inches;

            doc.Save(MyDir + @"\Artifacts\OdtSaveOptions.MeasureUnit.odt");
            //ExEnd
        }

        [Test]
        [TestCase(SaveFormat.Odt)]
        [TestCase(SaveFormat.Ott)]
        public void SaveDocumentEncryptedWithAPassword(SaveFormat saveFormat)
        {
            //ExStart
            //ExFor:OdtSaveOptions.Password
            //ExSummary:Shows how to encrypted your odt/ott documents with a password.
            Document doc = new Document(MyDir + "Document.docx");

            OdtSaveOptions saveOptions = new OdtSaveOptions(saveFormat);
            saveOptions.Password = "@sposeEncrypted_1145";

            // Saving document using password property of OdtSaveOptions
            doc.Save(MyDir + @"\Artifacts\OdtSaveOptions.SaveDocumentEncryptedWithAPassword" +
                     FileFormatUtil.SaveFormatToExtension(saveFormat), saveOptions);
            //ExEnd

            // Check that all documents are encrypted with a password
            FileFormatInfo docInfo = FileFormatUtil.DetectFileFormat(
                MyDir + @"\Artifacts\OdtSaveOptions.SaveDocumentEncryptedWithAPassword" +
                FileFormatUtil.SaveFormatToExtension(saveFormat));
            Assert.IsTrue(docInfo.IsEncrypted);
        }

        [Test]
        [TestCase(SaveFormat.Odt)]
        [TestCase(SaveFormat.Ott)]
        public void WorkWithDocumentEncryptedWithAPassword(SaveFormat saveFormat)
        {
            //ExStart
            //ExFor:OdtSaveOptions.#ctor(String)
            //ExSummary:Shows how to load and change odt/ott encrypted document
            Document doc = new Document(MyDir + "OdtSaveOptions.LoadDocumentEncryptedWithAPassword" +
                                        FileFormatUtil.SaveFormatToExtension(saveFormat),
                new LoadOptions("@sposeEncrypted_1145"));

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveToDocumentEnd();
            builder.Writeln("Encrypted document after changes.");

            // Saving document using new instance of OdtSaveOptions
            doc.Save(MyDir + @"\Artifacts\OdtSaveOptions.LoadDocumentEncryptedWithAPassword" +
                     FileFormatUtil.SaveFormatToExtension(saveFormat), new OdtSaveOptions("@sposeEncrypted_1145"));
            //ExEnd

            // Check that document is still encrypted with a password
            FileFormatInfo docInfo = FileFormatUtil.DetectFileFormat(
                MyDir + @"\Artifacts\OdtSaveOptions.LoadDocumentEncryptedWithAPassword" +
                FileFormatUtil.SaveFormatToExtension(saveFormat));
            Assert.IsTrue(docInfo.IsEncrypted);
        }
    }
}