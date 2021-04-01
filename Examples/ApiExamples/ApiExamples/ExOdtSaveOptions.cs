// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    internal class ExOdtSaveOptions : ApiExampleBase
    {
        [TestCase(false)]
        [TestCase(true)]
        public void Odt11Schema(bool exportToOdt11Specs)
        {
            //ExStart
            //ExFor:OdtSaveOptions
            //ExFor:OdtSaveOptions.#ctor
            //ExFor:OdtSaveOptions.IsStrictSchema11
            //ExSummary:Shows how to make a saved document conform to an older ODT schema.
            Document doc = new Document(MyDir + "Rendering.docx");

            OdtSaveOptions saveOptions = new OdtSaveOptions
            {
                MeasureUnit = OdtSaveMeasureUnit.Centimeters,
                IsStrictSchema11 = exportToOdt11Specs
            };

            doc.Save(ArtifactsDir + "OdtSaveOptions.Odt11Schema.odt", saveOptions);
            //ExEnd
            
            doc = new Document(ArtifactsDir + "OdtSaveOptions.Odt11Schema.odt");

            Assert.AreEqual(Aspose.Words.MeasurementUnits.Centimeters, doc.LayoutOptions.RevisionOptions.MeasurementUnit);

            if (exportToOdt11Specs)
            {
                Assert.AreEqual(2, doc.Range.FormFields.Count);
                Assert.AreEqual(FieldType.FieldFormTextInput, doc.Range.FormFields[0].Type);
                Assert.AreEqual(FieldType.FieldFormCheckBox, doc.Range.FormFields[1].Type);
            }
            else
            {
                Assert.AreEqual(3, doc.Range.FormFields.Count);
                Assert.AreEqual(FieldType.FieldFormTextInput, doc.Range.FormFields[0].Type);
                Assert.AreEqual(FieldType.FieldFormCheckBox, doc.Range.FormFields[1].Type);
                Assert.AreEqual(FieldType.FieldFormDropDown, doc.Range.FormFields[2].Type);
            }
        }

        [TestCase(OdtSaveMeasureUnit.Centimeters)]
        [TestCase(OdtSaveMeasureUnit.Inches)]
        public void MeasurementUnits(OdtSaveMeasureUnit odtSaveMeasureUnit)
        {
            //ExStart
            //ExFor:OdtSaveOptions
            //ExFor:OdtSaveOptions.MeasureUnit
            //ExFor:OdtSaveMeasureUnit
            //ExSummary:Shows how to use different measurement units to define style parameters of a saved ODT document.
            Document doc = new Document(MyDir + "Rendering.docx");

            // When we export the document to .odt, we can use an OdtSaveOptions object to modify how we save the document.
            // We can set the "MeasureUnit" property to "OdtSaveMeasureUnit.Centimeters"
            // to define content such as style parameters using the metric system, which Open Office uses. 
            // We can set the "MeasureUnit" property to "OdtSaveMeasureUnit.Inches"
            // to define content such as style parameters using the imperial system, which Microsoft Word uses.
            OdtSaveOptions saveOptions = new OdtSaveOptions
            {
                MeasureUnit = odtSaveMeasureUnit
            };

            doc.Save(ArtifactsDir + "OdtSaveOptions.Odt11Schema.odt", saveOptions);
            //ExEnd

            switch (odtSaveMeasureUnit)
            {
                case OdtSaveMeasureUnit.Centimeters:
                    TestUtil.DocPackageFileContainsString("<style:paragraph-properties fo:orphans=\"2\" fo:widows=\"2\" style:tab-stop-distance=\"1.27cm\" />",
                        ArtifactsDir + "OdtSaveOptions.Odt11Schema.odt", "styles.xml");
                    break;
                case OdtSaveMeasureUnit.Inches:
                    TestUtil.DocPackageFileContainsString("<style:paragraph-properties fo:orphans=\"2\" fo:widows=\"2\" style:tab-stop-distance=\"0.5in\" />",
                        ArtifactsDir + "OdtSaveOptions.Odt11Schema.odt", "styles.xml");
                    break;
            }
        }

        [TestCase(SaveFormat.Odt)]
        [TestCase(SaveFormat.Ott)]
        public void Encrypt(SaveFormat saveFormat)
        {
            //ExStart
            //ExFor:OdtSaveOptions.#ctor(SaveFormat)
            //ExFor:OdtSaveOptions.Password
            //ExFor:OdtSaveOptions.SaveFormat
            //ExSummary:Shows how to encrypt a saved ODT/OTT document with a password, and then load it using Aspose.Words.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            // Create a new OdtSaveOptions, and pass either "SaveFormat.Odt",
            // or "SaveFormat.Ott" as the format to save the document in. 
            OdtSaveOptions saveOptions = new OdtSaveOptions(saveFormat);
            saveOptions.Password = "@sposeEncrypted_1145";

            string extensionString = FileFormatUtil.SaveFormatToExtension(saveFormat);

            // If we open this document with an appropriate editor,
            // it will prompt us for the password we specified in the SaveOptions object.
            doc.Save(ArtifactsDir + "OdtSaveOptions.Encrypt" + extensionString, saveOptions);

            FileFormatInfo docInfo = FileFormatUtil.DetectFileFormat(ArtifactsDir + "OdtSaveOptions.Encrypt" + extensionString);

            Assert.IsTrue(docInfo.IsEncrypted);

            // If we wish to open or edit this document again using Aspose.Words,
            // we will have to provide a LoadOptions object with the correct password to the loading constructor.
            doc = new Document(ArtifactsDir + "OdtSaveOptions.Encrypt" + extensionString,
                new LoadOptions("@sposeEncrypted_1145"));

            Assert.AreEqual("Hello world!", doc.GetText().Trim());
            //ExEnd
        }
    }
}