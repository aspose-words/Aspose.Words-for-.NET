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
    public class OpenAndAddTextToWordDocument : TestUtil
    {
        [Test]
        public void AddText()
        {
            using (WordprocessingDocument originalDocument = WordprocessingDocument.Open(MyDir + "Document.docx", false))
            {
                // Create a new Wordprocessing document.
                using (WordprocessingDocument newDocument = WordprocessingDocument.Create(ArtifactsDir + "Add text - OpenXML.docx", WordprocessingDocumentType.Document))
                {
                    // Add a main document part to the new document.
                    MainDocumentPart newMainPart = newDocument.AddMainDocumentPart();
                    newMainPart.Document = new Document(new Body());

                    // Copy content from the original document to the new document.
                    MainDocumentPart originalMainPart = originalDocument.MainDocumentPart;
                    newMainPart.Document.Body = (Body)originalMainPart.Document.Body.Clone();

                    Body body = newMainPart.Document.Body;

                    // Create a new paragraph with the text you want to add
                    Paragraph newParagraph = new Paragraph();
                    Run newRun = new Run();
                    Text newText = new Text("This is the text added to the end of the document.");
                    newRun.Append(newText);
                    newParagraph.Append(newRun);

                    // Append the new paragraph to the body
                    body.Append(newParagraph);

                    // Save the changes
                    newMainPart.Document.Save();
                }
            }
        }
    }
}
