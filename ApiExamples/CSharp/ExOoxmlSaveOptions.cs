// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

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
            //ExFor:OoxmlCompliance.Iso29500_2008_Strict
            //ExSummary:Shows conversion vml shapes to dml using Iso29500_2008_Strict option
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            //Set Word2003 version for document, for inserting image as vml shape
            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2003);
            
            Shape image = builder.InsertImage(MyDir + @"dotnet-logo.png");

            // Loop through all single shapes inside document.
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                Assert.AreEqual(ShapeMarkupLanguage.Vml, shape.MarkupLanguage);
            }

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Strict; //Iso29500_2008 does not allow vml shapes, so you need to use OoxmlCompliance.Iso29500_2008_Strict for converting vml to dml shapes
            saveOptions.SaveFormat = SaveFormat.Docx;

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, saveOptions);

            //Assert that image have drawingML markup language
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                Assert.AreEqual(ShapeMarkupLanguage.Dml, shape.MarkupLanguage);
            }
            //ExEnd
        }
    }
}
