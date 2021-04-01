// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class ChangeOrReplaceHeaderAndFooter : TestUtil
    {
        [Test]
        public void ChangeOrReplaceHeaderAndFooterFeature()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Aspose.Words Header");
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.Write("Aspose.Words Footer");

            doc.Save(ArtifactsDir + "Change or replace header and footer - Aspose.Words.docx");
        }
    }
}
