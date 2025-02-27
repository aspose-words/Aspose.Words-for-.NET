// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;
using System.IO;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class AppendText: TestUtil
    {
        [Test]
        public void AddTextOpenXml()
        {
            //ExStart:AddTextOpenXml
            //GistId:bab40e2c44b7e59094cc177a8d5204d3
            File.Copy(MyDir + "Document.docx", ArtifactsDir + "Add text - OpenXML.docx", true);

            using WordprocessingDocument doc = WordprocessingDocument.Open(ArtifactsDir + "Add text - OpenXML.docx", true);

            MainDocumentPart mainPart = doc.MainDocumentPart;
            Body body = doc.MainDocumentPart.Document.Body;

            // Create a new paragraph with the text you want to add
            Paragraph newParagraph = new Paragraph();
            Run newRun = new Run();
            Text newText = new Text("This is the text added to the end of the document.");
            newRun.Append(newText);
            newParagraph.Append(newRun);

            // Append the new paragraph to the body
            body.Append(newParagraph);

            // Save the changes
            mainPart.Document.Save();
            //ExEnd:AddTextOpenXml
        }
    }
}
