// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class OpenDocumentFromStream : TestUtil
    {
        [Test]
        public void OpenDocumentFromStreamFeature()
        {
            using (Stream stream = File.Open(MyDir + "Document.docx", FileMode.Open))
            {
                using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(stream, true))
                {
                    Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
                    Paragraph para = body.AppendChild(new Paragraph());

                    Run run = para.AppendChild(new Run());
                    run.AppendChild(new Text("Append text in body - Open and add to wordprocessing stream"));
                }
            }
        }
    }
}
