// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using Aspose.Words;
using Aspose.Words.Drawing;
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
            //ExSummary:Shows conversion VML shapes to DML using Iso29500_2008_Strict
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
    }
}