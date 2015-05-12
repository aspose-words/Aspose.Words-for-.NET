//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.Collections;

using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Markup;
using Aspose.Words.Fields;
using Aspose.Words.Lists;

namespace PageSplitterExample
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
            if (paragraph.IsListItem)
            {
                List paraList = paragraph.ListFormat.List;
                ListLevel currentLevel = paragraph.ListFormat.ListLevel;

                // Since we have encountered a list item we need to check if this will reset
                // any subsequent list levels and if so then update the numbering of the level.
                int currentListLevelNumber = paragraph.ListFormat.ListLevelNumber;
                for (int i = currentListLevelNumber + 1; i < paraList.ListLevels.Count; i++)
                {
                    ListLevel paraLevel = paraList.ListLevels[i];

                    if (paraLevel.RestartAfterLevel >= currentListLevelNumber)
                    {
                        // This list level needs to be reset after the current list number.
                        mListLevelToListNumberLookup[paraLevel] = paraLevel.StartAt;
                    }
                }

                // A list which was used on a previous page is present on a different page, the list
                // needs to be copied so list numbering is retained when extracting individual pages.
                if (ContainsListLevelAndPageChanged(paragraph))
                {
                    List copyList = paragraph.Document.Lists.AddCopy(paraList);
                    mListLevelToListNumberLookup[currentLevel] = paragraph.ListLabel.LabelValue;

                    // Set the numbering of each list level to start at the numbering of the level on the previous page.
                    for (int i = 0; i < paraList.ListLevels.Count; i++)
                    {
                        ListLevel paraLevel = paraList.ListLevels[i];

                        if (mListLevelToListNumberLookup.ContainsKey(paraLevel))
                            copyList.ListLevels[i].StartAt = (int)mListLevelToListNumberLookup[paraLevel];
                    }

                    mListToReplacementListLookup[paraList] = copyList;
                }

                if (mListToReplacementListLookup.ContainsKey(paraList))
                {
                    // This paragraph belongs to a list from a previous page. Apply the replacement list.
                    paragraph.ListFormat.List = (List)mListToReplacementListLookup[paraList];
                    // This is a trick to get the spacing of the list level to set correctly.
                    paragraph.ListFormat.ListLevelNumber += 0;
                }

                mListLevelToPageLookup[currentLevel] = mPageNumberFinder.GetPage(paragraph);
                mListLevelToListNumberLookup[currentLevel] = paragraph.ListLabel.LabelValue;
            }

            Section prevSection = (Section)paragraph.ParentSection.PreviousSibling;
            Paragraph prevBodyPara = paragraph.PreviousSibling as Paragraph;

            Paragraph prevSectionPara = prevSection != null && paragraph == paragraph.ParentSection.Body.FirstChild ? prevSection.Body.LastParagraph : null;
            Paragraph prevParagraph = prevBodyPara != null ? prevBodyPara : prevSectionPara;

            if (paragraph.IsEndOfSection && !paragraph.HasChildNodes)
                paragraph.Remove();

            // Paragraphs across pages can merge or remove spacing depending upon the previous paragraph.
            if (prevParagraph != null)
            {
                if (mPageNumberFinder.GetPage(paragraph) != mPageNumberFinder.GetPageEnd(prevParagraph))
                {
                    if (paragraph.IsListItem && prevParagraph.IsListItem && !prevParagraph.IsEndOfSection)
                       prevParagraph.ParagraphFormat.SpaceAfter = 0;
                    else if (prevParagraph.ParagraphFormat.StyleName == paragraph.ParagraphFormat.StyleName && paragraph.ParagraphFormat.NoSpaceBetweenParagraphsOfSameStyle)
                        paragraph.ParagraphFormat.SpaceBefore = 0;
                    else if (paragraph.ParagraphFormat.PageBreakBefore || (prevParagraph.IsEndOfSection && prevSection.PageSetup.SectionStart != SectionStart.NewColumn))
                        paragraph.ParagraphFormat.SpaceBefore = System.Math.Max(paragraph.ParagraphFormat.SpaceBefore - prevParagraph.ParagraphFormat.SpaceAfter, 0);
                    else
                        paragraph.ParagraphFormat.SpaceBefore = 0;
                }
            }

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitSectionStart(Section section)
        {
            mSectionCount++;
            Section previousSection = (Section)section.PreviousSibling;

            // If there is a previous section attempt to copy any linked header footers otherwise they will not appear in an 
            // extracted document if the previous section is missing.
            if (previousSection != null)
            {
                if (!section.PageSetup.RestartPageNumbering)
                {
                    section.PageSetup.RestartPageNumbering = true;
                    section.PageSetup.PageStartingNumber = previousSection.PageSetup.PageStartingNumber + mPageNumberFinder.PageSpan(previousSection);
                }

                foreach (HeaderFooter previousHeaderFooter in previousSection.HeadersFooters)
                {
                    if (section.HeadersFooters[previousHeaderFooter.HeaderFooterType] == null)
                    {
                        HeaderFooter newHeaderFooter = (HeaderFooter)previousSection.HeadersFooters[previousHeaderFooter.HeaderFooterType].Clone(true);
                        section.HeadersFooters.Add(newHeaderFooter);
                    }
                }
            }

            // Manually set the result of these fields before sections are split.
            foreach (HeaderFooter headerFooter in section.HeadersFooters)
            {
                foreach (Field field in headerFooter.Range.Fields)
                {
                    if (field.Type == FieldType.FieldSection || field.Type == FieldType.FieldSectionPages)
                    {
                        field.Result = (field.Type == FieldType.FieldSection) ? mSectionCount.ToString() :
                            mPageNumberFinder.PageSpan(section).ToString();
                        field.IsLocked = true;
                    }
                }
            }

            // All fields in the body should stay the same, this also improves field update time.
            foreach (Field field in section.Body.Range.Fields)
                field.IsLocked = true;

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitDocumentEnd(Document doc)
        {
            // All sections have separate headers and footers now, update the fields in all headers and footers
            // to the correct values. This allows each page to maintain the correct field results even when
            // PAGE or IF fields are used.
            doc.UpdateFields();

            foreach (HeaderFooter headerFooter in doc.GetChildNodes(NodeType.HeaderFooter, true))
            {
                foreach (Field field in headerFooter.Range.Fields)
                    field.IsLocked = true;
            }

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitSmartTagEnd(SmartTag smartTag)
        {
            if(IsCompositeAcrossPage(smartTag))
                SplitComposite(smartTag);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitCustomXmlMarkupEnd(CustomXmlMarkup customXmlMarkup)
        {
            if (IsCompositeAcrossPage(customXmlMarkup))
                SplitComposite(customXmlMarkup);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitStructuredDocumentTagEnd(StructuredDocumentTag sdt)
        {
            if (IsCompositeAcrossPage(sdt)) 
                SplitComposite(sdt);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitCellEnd(Cell cell)
        {
            if (IsCompositeAcrossPage(cell)) 
                SplitComposite(cell);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitRowEnd(Row row)
        {
            if (IsCompositeAcrossPage(row)) 
                SplitComposite(row);

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitTableEnd(Table table)
        {
            if (IsCompositeAcrossPage(table))
            {
                // Copy any header rows to other pages.
                Stack stack = new Stack(table.Rows.ToArray());

                foreach (Table cloneTable in SplitComposite(table))
                {
                    foreach (Row row in stack)
                    {
                        if (row.RowFormat.HeadingFormat)
                            cloneTable.PrependChild(row.Clone(true));
                    }
                }
            }

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitParagraphEnd(Paragraph paragraph)
        {
            if (IsCompositeAcrossPage(paragraph))
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

                    // Reset spacing of split paragraphs as additional spacing is removed.
                    clonePara.ParagraphFormat.SpaceBefore = 0;
                    paragraph.ParagraphFormat.SpaceAfter = 0;
                }
            }

            return VisitorAction.Continue;
        }

        public override VisitorAction VisitSectionEnd(Section section)
        {
            if (IsCompositeAcrossPage(section))
            {
                // If a TOC field spans across more than one page then the hyperlink formatting may show through.
                // Remove direct formatting to avoid this.
                foreach (FieldStart start in section.GetChildNodes(NodeType.FieldStart, true))
                {
                    if (start.FieldType == FieldType.FieldTOC)
                    {
                        Field field = start.GetField();
                        Node node = field.Separator;

                        while ((node = node.NextPreOrder(section)) != field.End)
                            if (node.NodeType == NodeType.Run)
                                ((Run)node).Font.ClearFormatting();
                    }
                }

                foreach (Section cloneSection in SplitComposite(section))
                {
                    cloneSection.PageSetup.SectionStart = SectionStart.NewPage;
                    cloneSection.PageSetup.RestartPageNumbering = true;
                    cloneSection.PageSetup.PageStartingNumber = section.PageSetup.PageStartingNumber + (section.Document.IndexOf(cloneSection) - section.Document.IndexOf(section));
                    cloneSection.PageSetup.DifferentFirstPageHeaderFooter = false;

                    RemovePageBreaksFromParagraph(cloneSection.Body.LastParagraph);
                }

                RemovePageBreaksFromParagraph(section.Body.LastParagraph);

                // Add new page numbering for the body of the section as well.
                mPageNumberFinder.AddPageNumbersForNode(section.Body, mPageNumberFinder.GetPage(section), mPageNumberFinder.GetPageEnd(section));
            }

            return VisitorAction.Continue;
        }

        private bool IsCompositeAcrossPage(CompositeNode composite)
        {
            return mPageNumberFinder.PageSpan(composite) > 1;
        }

        private bool ContainsListLevelAndPageChanged(Paragraph para)
        {
            return mListLevelToPageLookup.ContainsKey(para.ListFormat.ListLevel) && (int)mListLevelToPageLookup[para.ListFormat.ListLevel] != mPageNumberFinder.GetPage(para);
        }

        private void RemovePageBreaksFromParagraph(Paragraph para)
        {
            if (para != null)
            {
                foreach (Run run in para.Runs)
                    run.Text = run.Text.Replace(ControlChar.PageBreak, string.Empty);
            }
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

        private Hashtable mListLevelToListNumberLookup = new Hashtable();
        private Hashtable mListToReplacementListLookup = new Hashtable();
        private Hashtable mListLevelToPageLookup = new Hashtable();
        private PageNumberFinder mPageNumberFinder;
        private int mSectionCount;
    }
}