// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class CreateADocument : TestUtil
    {
        [Test]
        public void CreateADocumentFeature()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello World!");

            doc.Save(ArtifactsDir + "Create a document - Aspose.Words.docx");
        }
    }
}
