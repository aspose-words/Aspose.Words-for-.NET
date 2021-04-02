// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features.Features_missing_in_OpenXML
{
    [TestFixture]
    public class ConvertToRtf : TestUtil
    {
        [Test]
        public static void ConvertToRtfFeature()
        {
            Document doc = new Document(MyDir + "Document.docx");
            
            doc.Save(ArtifactsDir + "Convert to Rtf - Aspose.Words.rtf");
        }
    }
}
