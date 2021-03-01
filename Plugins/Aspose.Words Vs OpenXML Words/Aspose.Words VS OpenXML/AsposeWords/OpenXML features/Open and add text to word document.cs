// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class OpenAndAddTextToWordDocument : TestUtil
    {
        [Test]
        public void OpenAndAddTextToWordDocumentFeature()
        {
            // Open a WordprocessingDocument for editing using the filepath.
            using (WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(MyDir + "Document.docx", true))
            {
                // Assign a reference to the existing document body.
                Body body = wordprocessingDocument.MainDocumentPart.Document.Body;

                // Add new text.
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text("Append text in body - Open and add text to word document"));
            }
        }
    }
}
