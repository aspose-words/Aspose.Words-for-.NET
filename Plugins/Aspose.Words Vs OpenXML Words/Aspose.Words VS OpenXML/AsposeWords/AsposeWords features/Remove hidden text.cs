// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class RemoveHiddenText : TestUtil
    {
        [Test]
        public void RemoveHiddenTextFeature()
        {
            Document doc = new Document(MyDir + "Remove hidden text.docx");
            
            foreach (Paragraph par in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                par.ParagraphBreakFont.Hidden = false;
                foreach (Run run in par.GetChildNodes(NodeType.Run, true))
                {
                    if (run.Font.Hidden)
                        run.Font.Hidden = false;
                }
            }

            doc.Save(ArtifactsDir + "Remove hidden text - Aspose.Words.docx");
        }
    }
}
