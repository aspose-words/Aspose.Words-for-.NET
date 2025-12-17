// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class RemovePageBreaks : TestUtil
    {
        [Test]
        public void RemovePageBreaksAsposeWords()
        {
            //ExStart:RemovePageBreaksAsposeWords
            //GistId:8cceaff98abfaa9643169bb00de01e4b
            Document doc = new Document(MyDir + "Remove page breaks.docx");

            // Remove all page breaks.
            foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                foreach (Run run in paragraph.Runs)
                {
                    if (run.Text.Contains(ControlChar.PageBreak))
                        run.Text = run.Text.Replace(ControlChar.PageBreak.ToString(), string.Empty);
                }
            }

            doc.Save(ArtifactsDir + "Remove page breaks - Aspose.Words.docx");
            //ExEnd:RemovePageBreaksAsposeWords
        }
    }
}

