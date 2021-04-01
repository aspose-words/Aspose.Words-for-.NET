// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.OpenXML_features
{
    [TestFixture]
    public class RemoveSectionBreaks : TestUtil
    {
        [Test]
        public void RemoveSectionBreaksFeature()
        {
            using (WordprocessingDocument myDoc = WordprocessingDocument.Open(MyDir + "Remove section breaks.docx", true))
            {
                MainDocumentPart mainPart = myDoc.MainDocumentPart;
                List<ParagraphProperties> paraProps = mainPart.Document.Descendants<ParagraphProperties>()
                    .Where(IsSectionProps).ToList();

                foreach (ParagraphProperties pPr in paraProps)
                    pPr.RemoveChild(pPr.GetFirstChild<SectionProperties>());

                using (Stream stream = File.Create(ArtifactsDir + "Remove section breaks - OpenXML.docx"))
                {
                    mainPart.Document.Save(stream);
                }
            }
        }

        private static bool IsSectionProps(ParagraphProperties pPr)
        {
            SectionProperties sectPr = pPr.GetFirstChild<SectionProperties>();

            if (sectPr == null)
                return false;

            return true;
        }
    }
}
