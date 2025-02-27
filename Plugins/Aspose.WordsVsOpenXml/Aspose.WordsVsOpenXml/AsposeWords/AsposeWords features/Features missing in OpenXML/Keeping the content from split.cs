// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

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
                para.ParagraphFormat.KeepWithNext = true;

            dstDoc.AppendDocument(srcDoc, ImportFormatMode.KeepSourceFormatting);
            dstDoc.Save(ArtifactsDir + "Keeping the content from split - Aspose.Words.docx");
        }
    }
}
