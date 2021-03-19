// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System.IO;
using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class OpenDocumentFromStream : TestUtil
    {
        [Test]
        public void OpenDocumentFromStreamFeature()
        {
            Stream stream = File.Open(MyDir + "Document.docx", FileMode.Open);

            using (stream)
            {
                Document doc = new Document(stream);
                DocumentBuilder builder = new DocumentBuilder(doc);

                builder.Writeln("Append text in body - Open and add to wordprocessing stream");
                
                doc.Save(ArtifactsDir + "Open document from stream - Aspose.Words.docx");
            }
        }
    }
}
