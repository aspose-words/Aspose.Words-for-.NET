// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class ConvertFromDocmToDocx : TestUtil
    {
        [Test]
        public void ConvertFromDocmToDocxFeature()
        {
            Document doc = new Document(MyDir + "Convert from docm to docx.docm");
            doc.Save(ArtifactsDir + "Convert from docm to docx - Aspose.Words.docx", SaveFormat.Docx);
        }
    }
}
