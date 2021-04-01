// Copyright (c) Aspose 2002-2021. All Rights Reserved.

using Aspose.Words;
using NUnit.Framework;

namespace AsposeWordsVSOpenXML.AsposeWords_features
{
    [TestFixture]
    public class RemoveSectionBreaks : TestUtil
    {
        [Test]
        public void RemoveAllSectionBreaks()
        {
            Document doc = new Document(MyDir + "Remove section breaks.docx");

            // Loop through all sections starting from the section that precedes the last one 
            // and moving to the first section.
            for (int i = doc.Sections.Count - 2; i >= 0; i--)
            {
                // Copy the content of the current section to the beginning of the last section.
                doc.LastSection.PrependContent(doc.Sections[i]);

                // Remove the copied section.
                doc.Sections[i].Remove();
            }
        }
    }
}
