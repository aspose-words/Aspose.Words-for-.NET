// Copyright (c) 2001-2017 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
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
        public void Iso29500Strict()
        {
            //ExStart
            //ExFor:OoxmlCompliance
            //ExFor:OoxmlSaveOptions.Compliance
            //ExSummary:Shows conversion VML shapes to DML using ISO/IEC 29500:2008 Strict compliance level
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            //Set Word2003 version for document, for inserting image as VML shape
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2003);

            builder.InsertImage(ImageDir + "dotnet-logo.png");

            // Loop through all single shapes inside document.
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                Console.WriteLine(shape.MarkupLanguage);
                Assert.AreEqual(ShapeMarkupLanguage.Vml, shape.MarkupLanguage);//ExSkip
            }

            //Iso29500_2008 does not allow VML shapes, so you need to use OoxmlCompliance.Iso29500_2008_Strict for converting VML to DML shapes
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Strict; 
            saveOptions.SaveFormat = SaveFormat.Docx;
            //ExEnd

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, saveOptions);

            //Assert that image have drawingML markup language
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                Assert.AreEqual(ShapeMarkupLanguage.Dml, shape.MarkupLanguage);
            }
        }

        [Test]
        public void RestartingDocumentList()
        {
            //ExStart
            //ExFor:List.IsRestartAtEachSection
            //ExSummary:Shows how to specify that the list has to be restarted at each section.
            Document doc = new Document();

            doc.Lists.Add(ListTemplate.NumberDefault);

            Aspose.Words.Lists.List list = doc.Lists[0];

            // Set true to specify that the list has to be restarted at each section.
            list.IsRestartAtEachSection = true;

            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.ListFormat.List = list;

            for (int i = 1; i <= 45; i++)
            {
                builder.Write($"List Item {i}\n");

                // Insert section break.
                if (i == 15 || i == 30)
                    builder.InsertBreak(BreakType.SectionBreakNewPage);
            }

            // IsRestartAtEachSection will be written only if compliance is higher then OoxmlComplianceCore.Ecma376
            OoxmlSaveOptions options = new OoxmlSaveOptions();
            options.Compliance = OoxmlCompliance.Iso29500_2008_Transitional;

            doc.Save(MyDir + @"\Artifacts\RestartingDocumentList.docx", options);
            //ExEnd
        }

        [Test]
        public void UpdatingLastSavedTimeDocument()
        {
            //ExStart
            //ExFor:SaveOptions.UpdateLastSavedTimeProperty
            //ExSummary:Shows how to update a document time property when you want to save it
            Document doc = new Document(MyDir + "Document.doc");

            //Get last saved time
            DateTime documentTimeBeforeSave = doc.BuiltInDocumentProperties.LastSavedTime;

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.UpdateLastSavedTimeProperty = true;
            //ExEnd

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, saveOptions);

            DateTime documentTimeAfterSave = doc.BuiltInDocumentProperties.LastSavedTime;

            Assert.AreNotEqual(documentTimeBeforeSave, documentTimeAfterSave);
        }
    }
}