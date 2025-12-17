using Aspose.Words;
using Aspose.Words.Fields;
using NUnit.Framework;
using System;

namespace DocsExamples.Programming_with_Documents.Contents_Management
{
    class WorkingWithTableOfContent : DocsExamplesBase
    {
        [Test]
        public void ChangeStyleOfTocLevel()
        {
            //ExStart:ChangeStyleOfTocLevel
            //GistId:db118a3e1559b9c88355356df9d7ea10
            Document doc = new Document();
            // Retrieve the style used for the first level of the TOC and change the formatting of the style.
            doc.Styles[StyleIdentifier.Toc1].Font.Bold = true;
            //ExEnd:ChangeStyleOfTocLevel
        }

        [Test]
        public void ChangeTocTabStops()
        {
            //ExStart:ChangeTocTabStops
            //GistId:db118a3e1559b9c88355356df9d7ea10
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
            //ExEnd:ChangeTocTabStops
        }

        [Test]
        public void ExtractToc()
        {
            //ExStart:ExtractToc
            //GistId:db118a3e1559b9c88355356df9d7ea10
            Document doc = new Document(MyDir + "Table of contents.docx");

            foreach (Field field in doc.Range.Fields)
            {
                if (field.Type.Equals(FieldType.FieldHyperlink))
                {
                    FieldHyperlink hyperlink = (FieldHyperlink)field;
                    if (hyperlink.SubAddress != null && hyperlink.SubAddress.StartsWith("_Toc"))
                    {
                        Paragraph tocItem = (Paragraph)field.Start.GetAncestor(NodeType.Paragraph);
                        Console.WriteLine(tocItem.ToString(SaveFormat.Text).Trim());
                        Console.WriteLine("------------------");
                        if (tocItem != null)
                        {
                            Bookmark bm = doc.Range.Bookmarks[hyperlink.SubAddress];
                            Paragraph pointer = (Paragraph)bm.BookmarkStart.GetAncestor(NodeType.Paragraph);
                            Console.WriteLine(pointer.ToString(SaveFormat.Text));
                        }
                    }
                }
            }
            //ExEnd:ExtractToc
        }
    }
}