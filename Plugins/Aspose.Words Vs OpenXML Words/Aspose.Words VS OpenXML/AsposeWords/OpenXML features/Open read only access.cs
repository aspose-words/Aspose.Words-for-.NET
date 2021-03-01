// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class OpenReadOnlyAccess : TestUtil
    {
        [Test]
        public void OpenReadOnlyAccessFeature()
        {
            // Open a WordprocessingDocument based on a filepath.
            using (WordprocessingDocument wordDocument =
                WordprocessingDocument.Open(MyDir + "Open readonly access.docx", false))
            {
                // Assign a reference to the existing document body.  
                Body body = wordDocument.MainDocumentPart.Document.Body;

                // Attempt to add some text.
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text("Append text in body, but text is not saved - Open wordprocessing document readonly"));

                // Call the "Save" method to generate an exception and show that access is read-only.
                using (Stream stream = File.Create(ArtifactsDir + "Open readonly access - OpenXML.docx"))
                {
                    wordDocument.MainDocumentPart.Document.Save(stream);
                }
            }
        }
    }
}
