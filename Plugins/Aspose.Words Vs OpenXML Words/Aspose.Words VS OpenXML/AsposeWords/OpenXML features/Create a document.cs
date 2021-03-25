// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class CreateADocument : TestUtil
    {
        [Test]
        public void CreateADocumentFeature()
        {
            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Create(ArtifactsDir + "Create a document - OpenXML.docx",
                    WordprocessingDocumentType.Document))
            {
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();
                // Create the document structure and add some text.
                mainPart.Document = new Document();
                
                Body body = mainPart.Document.AppendChild(new Body());
                Paragraph para = body.AppendChild(new Paragraph());
                
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text("Create text in body - Create wordprocessing document"));
            }
        }
    }
}
