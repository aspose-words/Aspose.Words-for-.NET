using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Contents_Managment
{
    class WorkingWithTableOfContent : DocsExamplesBase
    {
        [Test]
        public void ChangeStyleOfTocLevel()
        {
            //ExStart:ChangeStyleOfTOCLevel
            Document doc = new Document();
            // Retrieve the style used for the first level of the TOC and change the formatting of the style.
            doc.Styles[StyleIdentifier.Toc1].Font.Bold = true;
            //ExEnd:ChangeStyleOfTOCLevel
        }

        [Test]
        public void ChangeTocTabStops()
        {
            //ExStart:ChangeTOCTabStops
            Document doc = new Document(MyDir + "Table of contents.docx");

            foreach (Paragraph para in doc.GetChildNodes(NodeType.Paragraph, true))
            {
                // Check if this paragraph is formatted using the TOC result based styles.
                // This is any style between TOC and TOC9.
                if (para.ParagraphFormat.Style.StyleIdentifier >= StyleIdentifier.Toc1 &&
                    para.ParagraphFormat.Style.StyleIdentifier <= StyleIdentifier.Toc9)
                {
                    // Get the first tab used in this paragraph, this should be the tab used to align the page numbers.
                    TabStop tab = para.ParagraphFormat.TabStops[0];
                    
                    // Remove the old tab from the collection.
                    para.ParagraphFormat.TabStops.RemoveByPosition(tab.Position);
                    
                    // Insert a new tab using the same properties but at a modified position.
                    // We could also change the separators used (dots) by passing a different Leader type.
                    para.ParagraphFormat.TabStops.Add(tab.Position - 50, tab.Alignment, tab.Leader);
                }
            }

            doc.Save(ArtifactsDir + "WorkingWithTableOfContent.ChangeTocTabStops.docx");
            //ExEnd:ChangeTOCTabStops
        }
    }
}