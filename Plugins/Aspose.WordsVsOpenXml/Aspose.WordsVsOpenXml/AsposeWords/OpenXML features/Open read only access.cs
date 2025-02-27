// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

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
        public void OpenReadOnlyOpenXml()
        {
            //ExStart:OpenReadOnlyOpenXml
            //GistId:702c287894827f3d4ddd2ca4b170ed45
            using (var fileStream = new FileStream(MyDir + "Open readonly access.docx", FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using WordprocessingDocument doc = WordprocessingDocument.Open(fileStream, false);
                {
                    // Assign a reference to the existing document body.
                    Body body = doc.MainDocumentPart.Document.Body;

                    // Attempt to add some text.
                    Paragraph para = body.AppendChild(new Paragraph());
                    Run run = para.AppendChild(new Run());
                    run.AppendChild(new Text("This is the text added to the end of the document."));

                    // Call method to generate an exception and show that access is read-only.
                    using (Stream stream = File.Create(ArtifactsDir + "Open readonly access - OpenXML.docx"))
                        Assert.Throws(typeof(FileFormatException), () => doc.MainDocumentPart.Document.Save(stream));
                }
            }
            //ExEnd:OpenReadOnlyOpenXml
        }
    }
}
