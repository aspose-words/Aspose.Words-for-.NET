// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    class OpenReadOnlyAccess : TestUtil
    {
        [Test]
        public void OpenReadOnlyAccessFeature()
        {
            Document doc = new Document(MyDir + "Open ReadOnly access.docx", new LoadOptions("1234"));
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Append text in body - Open ReadOnly access");
            
            doc.Save(ArtifactsDir + "Open ReadOnly access - Aspose.Words.docx");
        }
    }
}
