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
    public class RemoveHiddenText : TestUtil
    {
        [Test]
        public void RemoveHiddenTextAsposeWords()
        {
            //ExStart:RemoveHiddenTextAsposeWords
            //GistDesc:Remove hidden text from document using C#
            Document doc = new Document(MyDir + "Remove hidden text.docx");
            
            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                para.ParagraphBreakFont.Hidden = false;
                foreach (Run run in para.GetChildNodes(NodeType.Run, true))
                {
                    if (run.Font.Hidden)
                        run.Font.Hidden = false;
                }
            }

            doc.Save(ArtifactsDir + "Remove hidden text - Aspose.Words.docx");
            //ExEnd:RemoveHiddenTextAsposeWords
        }
    }
}
