// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class RemoveHeaderFooter : TestUtil
    {
        [Test]
        public void RemoveHeaderFooterFeature()
        {
            Document doc = new Document(MyDir + "Document.docx");
            foreach (Section section in doc)
            {
                section.HeadersFooters.RemoveAt(0);

                // Odd pages use the primary footer.
                HeaderFooter footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];

                footer?.Remove();
            }

            doc.Save(ArtifactsDir + "Remove header and footer - Aspose.Words.docx");
        }
    }
}
