// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

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
        public void CreateNewDocumentOpenXml()
        {
            //ExStart:CreateNewDocumentOpenXml
            //GistId:e75459ad5b9ea7ac4cbea10ab631a491
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(ArtifactsDir + "Create new document - OpenXML.docx", 
                WordprocessingDocumentType.Document))
            {
                // Add a main document part.
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = new Body();
                Paragraph paragraph = new Paragraph();
                Run run = new Run();
                run.Append(new Text("Hello, Open XML!"));
                paragraph.Append(run);
                body.Append(paragraph);
                mainPart.Document.Append(body);
                mainPart.Document.Save();
            }
            //ExEnd:CreateNewDocumentOpenXml
        }
    }
}
