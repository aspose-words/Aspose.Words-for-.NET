//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.Collections;

using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Markup;

namespace PageSplitter
{
    /// <summary>
    /// Splits a document into multiple sections so that each page begins and ends at a section boundary.
    /// </summary>
    public class SectionSplitter : DocumentVisitor
    {
        public SectionSplitter(PageNumberFinder pageNumberFinder)
        {
            mPageNumberFinder = pageNumberFinder;
        }

        public override VisitorAction VisitParagraphStart(Paragraph paragraph)
        {
            return ContinueIfCompositeAcrossPageElseSkip(paragraph);
        }

        public override VisitorAction VisitTableStart(Table table)
        {
            return ContinueIfCompositeAcrossPageElseSkip(table);
        }

        public override VisitorAction VisitRowStart(Row row)
        {
            return ContinueIfCompositeAcrossPageElseSkip(row);
        }

        public override VisitorAction VisitCellStart(Cell cell)
        {
            return ContinueIfCompositeAcrossPageElseSkip(cell);
        }

        public override VisitorAction VisitCustomXmlMarkupStart(CustomXmlMarkup customXmlMarkup)
        {
            return ContinueIfCompositeAcrossPageElseSkip(customXmlMarkup);
        }

        public override VisitorAction VisitStructuredDocumentTagStart(StructuredDocumentTag sdt)
        {
            return ContinueIfCompositeAcrossPageElseSkip(sdt);
        }

        public override VisitorAction VisitSmartTagStart(SmartTag smartTag)
        {
            return ContinueIfCompositeAcrossPageElseSkip(smartTag);
        }

        public override VisitorAction VisitSectionStart(Section section)
        {
            Section previousSection = (Section)section.PreviousSibling;

            // If there is a previous section attempt to copy any linked header footers otherwise they will not appear in an 
            // extracted document if the previous section is missing.
            if (previousSection != null)
            {
                HeaderFooterCollection previousHeaderFooters = previousSection.HeadersFooters;

                if (!section.PageSetup.RestartPageNumbering)
                {
                    section.PageSetup.RestartPageNumbering = true;
                    section.PageSetup.PageStartingNumber = previousSection.PageSetup.PageStartingNumber + mPageNumberFinder.PageSpan(previousSection);
                }

                foreach (HeaderFooter previousHeaderFooter in previousHeaderFooters)
                {
                    if (section.HeadersFooters[previousHeaderFooter.HeaderFooterType] == null)
                    {
                        HeaderFooter newHeaderFooter = (HeaderFooter)previousHeaderFooters[previousHeaderFooter.HeaderFooterType].Clone(true);
                        section.HeadersFooters.Add(newHeaderFooter);
                    }
                }
            }

            return ContinueIfCompositeAcrossPageElseSkip(section);
        }

        public override VisitorAction VisitSmartTagEnd(SmartTag smartTag)
        {
            SplitComposite(smartTag);
            return VisitorAction.Continue;
        }

        public override VisitorAction VisitCustomXmlMarkupEnd(CustomXmlMarkup customXmlMarkup)
        {
            SplitComposite(customXmlMarkup);
            return VisitorAction.Continue;
        }

        public override VisitorAction VisitStructuredDocumentTagEnd(StructuredDocumentTag sdt)
        {
            SplitComposite(sdt);
            return VisitorAction.Continue;
        }

        public override VisitorAction VisitCellEnd(Cell cell)
        {
            SplitComposite(cell);
            return VisitorAction.Continue;
        }

        public override VisitorAction VisitRowEnd(Row row)
        {
            SplitComposite(row);
            return VisitorAction.Continue;
        }

        public override VisitorAction VisitTableEnd(Table table)
        {
            SplitComposite(table);
            return VisitorAction.Continue;
        }

        public override VisitorAction VisitParagraphEnd(Paragraph paragraph)
        {
            foreach (Paragraph clonePara in SplitComposite(paragraph))
            {
                // Remove list numbering from the cloned paragraph but leave the indent the same 
                // as the paragraph is supposed to be part of the item before.
                if (paragraph.IsListItem)
                {
                    double textPosition = clonePara.ListFormat.ListLevel.TextPosition;
                    clonePara.ListFormat.RemoveNumbers();
                    clonePara.ParagraphFormat.LeftIndent = textPosition;
                }

                // Reset spacing of split paragraphs in tables as additional spacing may cause them to look different.
                if (paragraph.IsInCell)
                {
                    clonePara.ParagraphFormat.SpaceBefore = 0;
                    paragraph.ParagraphFormat.SpaceAfter = 0;
                }
            }

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitSectionEnd(Section section)
        {
            foreach (Section cloneSection in SplitComposite(section))
            {
                cloneSection.PageSetup.SectionStart = SectionStart.NewPage;
                cloneSection.PageSetup.RestartPageNumbering = true;
                cloneSection.PageSetup.PageStartingNumber = section.PageSetup.PageStartingNumber + (section.Document.IndexOf(cloneSection) - section.Document.IndexOf(section));
                cloneSection.PageSetup.DifferentFirstPageHeaderFooter = false;
            }

            // Add new page numbering for the body of the section as well.
            mPageNumberFinder.AddPageNumbersForNode(section.Body, mPageNumberFinder.GetPage(section), mPageNumberFinder.GetPageEnd(section));
            
            return VisitorAction.Continue;
        }

        private VisitorAction ContinueIfCompositeAcrossPageElseSkip(CompositeNode composite)
        {
            return (mPageNumberFinder.PageSpan(composite) > 1) ? VisitorAction.Continue : VisitorAction.SkipThisNode;
        }

        private ArrayList SplitComposite(CompositeNode composite)
        {
            ArrayList splitNodes = new ArrayList();
            foreach (Node splitNode in FindChildSplitPositions(composite))
                splitNodes.Add(SplitCompositeAtNode(composite, splitNode));

            return splitNodes;
        }

        private ArrayList FindChildSplitPositions(CompositeNode node)
        {
            // A node may span across multiple pages so a list of split positions is returned.
            // The split node is the first node on the next page.
            ArrayList splitList = new ArrayList();

            int startingPage = mPageNumberFinder.GetPage(node);

            Node[] childNodes = node.NodeType == NodeType.Section ? 
                ((Section)node).Body.ChildNodes.ToArray() : node.ChildNodes.ToArray();

            foreach (Node childNode in childNodes)
            {
                int pageNum = mPageNumberFinder.GetPage(childNode);

                // If the page of the child node has changed then this is the split position. Add
                // this to the list.
                if (pageNum > startingPage)
                {
                    splitList.Add(childNode);
                    startingPage = pageNum;
                }

                if (mPageNumberFinder.PageSpan(childNode) > 1)
                    mPageNumberFinder.AddPageNumbersForNode(childNode, pageNum, pageNum);
            }

            // Split composites backward so the cloned nodes are inserted in the right order.
            splitList.Reverse();

            return splitList;
        }

        private CompositeNode SplitCompositeAtNode(CompositeNode baseNode, Node targetNode)
        {
            CompositeNode cloneNode = (CompositeNode)baseNode.Clone(false);

            Node node = targetNode;
            int currentPageNum = mPageNumberFinder.GetPage(baseNode);

            // Move all nodes found on the next page into the copied node. Handle row nodes separately.
            if (baseNode.NodeType != NodeType.Row)
            {
                CompositeNode composite = cloneNode;

                if (baseNode.NodeType == NodeType.Section)
                {
                    cloneNode = (CompositeNode)baseNode.Clone(true);
                    Section section = (Section)cloneNode;
                    section.Body.RemoveAllChildren();

                    composite = section.Body;
                }

                while (node != null)
                {
                    Node nextNode = node.NextSibling;
                    composite.AppendChild(node);
                    node = nextNode;
                }
            }
            else
            {
                // If we are dealing with a row then we need to add in dummy cells for the cloned row.
                int targetPageNum = mPageNumberFinder.GetPage(targetNode);
                Node[] childNodes = baseNode.ChildNodes.ToArray();

                foreach (Node childNode in childNodes)
                {
                    int pageNum = mPageNumberFinder.GetPage(childNode);

                    if (pageNum == targetPageNum)
                    {
                        cloneNode.LastChild.Remove();
                        cloneNode.AppendChild(childNode);
                    }
                    else if (pageNum == currentPageNum)
                    {
                        cloneNode.AppendChild(childNode.Clone(false));
                        if (cloneNode.LastChild.NodeType != NodeType.Cell)
                            ((CompositeNode)cloneNode.LastChild).AppendChild(((CompositeNode)childNode).FirstChild.Clone(false));
                    }
                }
            }

            // Insert the split node after the original.
            baseNode.ParentNode.InsertAfter(cloneNode, baseNode);

            // Update the new page numbers of the base node and the clone node including its descendents.
            // This will only be a single page as the cloned composite is split to be on one page.
            int currentEndPageNum = mPageNumberFinder.GetPageEnd(baseNode);
            mPageNumberFinder.AddPageNumbersForNode(baseNode, currentPageNum, currentEndPageNum - 1);
            mPageNumberFinder.AddPageNumbersForNode(cloneNode, currentEndPageNum, currentEndPageNum);

            foreach (Node childNode in cloneNode.GetChildNodes(NodeType.Any, true))
                mPageNumberFinder.AddPageNumbersForNode(childNode, currentEndPageNum, currentEndPageNum);

            return cloneNode;
        }

        private PageNumberFinder mPageNumberFinder;
    }
}