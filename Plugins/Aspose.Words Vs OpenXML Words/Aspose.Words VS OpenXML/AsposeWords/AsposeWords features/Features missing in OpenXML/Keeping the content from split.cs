// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features.Features_missing_in_OpenXML
{
    [TestFixture]
    public class KeepingTheContentFromSplit : TestUtil
    {
        [Test]
        public static void KeepingTheContentFromSplitFeature()
        {
            Document dstDoc = new Document(MyDir + "Joining mutiple documents 2.docx");
            Document srcDoc = new Document(MyDir + "Joining mutiple documents 1.docx");

            // Set the source document to appear straight after the destination document's content.
            srcDoc.FirstSection.PageSetup.SectionStart = SectionStart.Continuous;

            // Iterate through all sections in the source document.
            foreach (Paragraph para in srcDoc.GetChildNodes(NodeType.Paragraph, true))
            {
                para.ParagraphFormat.KeepWithNext = true;
            }

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            dstDoc.Save(ArtifactsDir + "Keeping the content from split - Aspose.Words.docx");
        }
    }
}
