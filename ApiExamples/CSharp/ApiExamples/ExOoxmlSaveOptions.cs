// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Lists;
using Aspose.Words.Saving;
using Aspose.Words.Settings;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    internal class ExOoxmlSaveOptions : ApiExampleBase
    {
        [Test]
        public void Password()
        {
            //ExStart
            //ExFor:OoxmlSaveOptions.Password
            //ExSummary:Shows how to create a password protected Office Open XML document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            // Create a SaveOptions object with a password and save our document with it
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.Password = "MyPassword";

            doc.Save(ArtifactsDir + "OoxmlSaveOptions.Password.docx", saveOptions);

            // This document cannot be opened like a normal document
            Assert.Throws<IncorrectPasswordException>(() => doc = new Document(ArtifactsDir + "OoxmlSaveOptions.Password.docx"));

            // We can open the document and access its contents by passing the correct password to a LoadOptions object
            doc = new Document(ArtifactsDir + "OoxmlSaveOptions.Password.docx", new LoadOptions("MyPassword"));

            Assert.AreEqual("Hello world!", doc.GetText().Trim());
            //ExEnd
        }

        [Test]
        public void Iso29500Strict()
        {
            //ExStart
            //ExFor:OoxmlSaveOptions
            //ExFor:OoxmlSaveOptions.#ctor
            //ExFor:OoxmlSaveOptions.SaveFormat
            //ExFor:OoxmlCompliance
            //ExFor:OoxmlSaveOptions.Compliance
            //ExFor:ShapeMarkupLanguage
            //ExSummary:Shows conversion VML shapes to DML using ISO/IEC 29500:2008 Strict compliance level.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Set Word2003 version for document, for inserting image as VML shape
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2003);
            builder.InsertImage(ImageDir + "Transparent background logo.png");

            Assert.AreEqual(ShapeMarkupLanguage.Vml, ((Shape)doc.GetChild(NodeType.Shape, 0, true)).MarkupLanguage);

            // Iso29500_2008 does not allow VML shapes
            // You need to use OoxmlCompliance.Iso29500_2008_Strict for converting VML to DML shapes
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
            {
                Compliance = OoxmlCompliance.Iso29500_2008_Strict,
                SaveFormat = SaveFormat.Docx
            };

            doc.Save(ArtifactsDir + "OoxmlSaveOptions.Iso29500Strict.docx", saveOptions);

            // The markup language of our shape has changed according to the compliance type 
            doc = new Document(ArtifactsDir + "OoxmlSaveOptions.Iso29500Strict.docx");
            
            Assert.AreEqual(ShapeMarkupLanguage.Dml, ((Shape)doc.GetChild(NodeType.Shape, 0, true)).MarkupLanguage);
            //ExEnd
        }

        [Test]
        [TestCase(false)]
        [TestCase(true)]
        public void RestartingDocumentList(bool doRestartListAtEachSection)
        {
            //ExStart
            //ExFor:List.IsRestartAtEachSection
            //ExSummary:Shows how to specify that the list has to be restarted at each section.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            doc.Lists.Add(ListTemplate.NumberDefault);

            Aspose.Words.Lists.List list = doc.Lists[0];

            // Set true to specify that the list has to be restarted at each section
            list.IsRestartAtEachSection = doRestartListAtEachSection;

            // IsRestartAtEachSection will be written only if compliance is higher then OoxmlComplianceCore.Ecma376
            OoxmlSaveOptions options = new OoxmlSaveOptions
            {
                Compliance = OoxmlCompliance.Iso29500_2008_Transitional
            };

            builder.ListFormat.List = list;

            builder.Writeln("List item 1");
            builder.Writeln("List item 2");
            builder.InsertBreak(BreakType.SectionBreakNewPage);
            builder.Writeln("List item 3");
            builder.Writeln("List item 4");

            doc.Save(ArtifactsDir + "OoxmlSaveOptions.RestartingDocumentList.docx", options);
            //ExEnd
            
            doc = new Document(ArtifactsDir + "OoxmlSaveOptions.RestartingDocumentList.docx");

            Assert.AreEqual(doRestartListAtEachSection, doc.Lists[0].IsRestartAtEachSection);
        }

        [Test]
        public void UpdatingLastSavedTimeDocument()
        {
            //ExStart
            //ExFor:SaveOptions.UpdateLastSavedTimeProperty
            //ExSummary:Shows how to update a document time property when you want to save it.
            Document doc = new Document(MyDir + "Document.docx");

            // Get last saved time
            DateTime documentTimeBeforeSave = doc.BuiltInDocumentProperties.LastSavedTime;

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
            {
                UpdateLastSavedTimeProperty = true
            };

            doc.Save(ArtifactsDir + "OoxmlSaveOptions.UpdatingLastSavedTimeDocument.docx", saveOptions);
            //ExEnd

            doc = DocumentHelper.SaveOpen(doc);
            DateTime documentTimeAfterSave = doc.BuiltInDocumentProperties.LastSavedTime;

            Assert.True(documentTimeBeforeSave < documentTimeAfterSave);
        }

        [Test]
        [TestCase(false)]
        [TestCase(true)]
        public void KeepLegacyControlChars(bool doKeepLegacyControlChars)
        {
            //ExStart
            //ExFor:OoxmlSaveOptions.KeepLegacyControlChars
            //ExFor:OoxmlSaveOptions.#ctor(SaveFormat)
            //ExSummary:Shows how to support legacy control characters when converting to .docx.
            Document doc = new Document(MyDir + "Legacy control character.doc");
 
            // Note that only one legacy character (ShortDateTime) is supported which declared in the "DOC" format
            OoxmlSaveOptions so = new OoxmlSaveOptions(SaveFormat.Docx);
            so.KeepLegacyControlChars = doKeepLegacyControlChars;
 
            doc.Save(ArtifactsDir + "OoxmlSaveOptions.KeepLegacyControlChars.docx", so);

            // Open the saved document and verify results
            doc = new Document(ArtifactsDir + "OoxmlSaveOptions.KeepLegacyControlChars.docx");

            if (doKeepLegacyControlChars)
                Assert.AreEqual("\u0013date \\@ \"M/d/yyyy\"\u0014\u0015\f", doc.FirstSection.Body.GetText());
            else
                Assert.AreEqual("\u001e\f", doc.FirstSection.Body.GetText());
            //ExEnd
        }
    }
}