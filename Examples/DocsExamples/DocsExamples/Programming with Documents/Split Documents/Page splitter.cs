using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Markup;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Split_Documents
{
    internal class PageSplitter : DocsExamplesBase
    {
        [Test]
        public void SplitDocuments()
        {
            SplitAllDocumentsToPages(MyDir);
        }

        public void SplitDocumentToPages(string docName)
        {
            string fileName = Path.GetFileNameWithoutExtension(docName);
            string extensionName = Path.GetExtension(docName);

            Console.WriteLine("Processing document: " + fileName + extensionName);

            Document doc = new Document(docName);

            // Split nodes in the document into separate pages.
            DocumentPageSplitter splitter = new DocumentPageSplitter(doc);

            // Save each page to the disk as a separate document.
            for (int page = 1; page <= doc.PageCount; page++)
            {
                Document pageDoc = splitter.GetDocumentOfPage(page);
                pageDoc.Save(Path.Combine(ArtifactsDir,
                    $"{fileName} - page{page} Out{extensionName}"));
            }
        }

        public void SplitAllDocumentsToPages(string folderName)
        {
            List<string> fileNames = Directory.GetFiles(folderName, "*.doc", SearchOption.TopDirectoryOnly)
                .Where(item => item.EndsWith(".doc")).ToList();

            foreach (string fileName in fileNames)
            {
                SplitDocumentToPages(fileName);
            }
        }
    }

    /// <summary>
    /// Splits a document into multiple documents, one per page.
    /// </summary>
    internal class DocumentPageSplitter
    {
        private readonly PageNumberFinder pageNumberFinder;

        /// <summary>
        /// Initializes a new instance of the <see cref="DocumentPageSplitter"/> class.
        /// This method splits the document into sections so that each page begins and ends at a section boundary.
        /// It is recommended not to modify the document afterwards.
        /// </summary>
        /// <param name="source">Source document</param>
        public DocumentPageSplitter(Document source)
        {
            pageNumberFinder = PageNumberFinderFactory.Create(source);
        }

        /// <summary>
        /// Gets the document this instance works with.
        /// </summary>
        private Document Document => pageNumberFinder.Document;

        /// <summary>
        /// Gets the document of a page.
        /// </summary>
        /// <param name="pageIndex">
        /// 1-based index of a page.
        /// </param>
        /// <returns>
        /// The <see cref="Document"/>.
        /// </returns>
        public Document GetDocumentOfPage(int pageIndex)
        {
            return GetDocumentOfPageRange(pageIndex, pageIndex);
        }

        /// <summary>
        /// Gets the document of a page range.
        /// </summary>
        /// <param name="startIndex">
        /// 1-based index of the start page.
        /// </param>
        /// <param name="endIndex">
        /// 1-based index of the end page.
        /// </param>
        /// <returns>
        /// The <see cref="Document"/>.
        /// </returns>
        public Document GetDocumentOfPageRange(int startIndex, int endIndex)
        {
            Document result = (Document) Document.Clone(false);
            foreach (Node section in pageNumberFinder.RetrieveAllNodesOnPages(startIndex, endIndex,
                NodeType.Section))
            {
                result.AppendChild(result.ImportNode(section, true));
            }

            return result;
        }
    }

    /// <summary>
    /// Provides methods for extracting nodes of a document which are rendered on a specified pages.
    /// </summary>
    public class PageNumberFinder
    {
        // Maps node to a start/end page numbers.
        // This is used to override baseline page numbers provided by the collector when the document is split.
        private readonly IDictionary<Node, int> nodeStartPageLookup = new Dictionary<Node, int>();
        private readonly IDictionary<Node, int> nodeEndPageLookup = new Dictionary<Node, int>();
        private readonly LayoutCollector collector;

        // Maps page number to a list of nodes found on that page.
        private IDictionary<int, IList<Node>> reversePageLookup;

        /// <summary>
        /// Initializes a new instance of the <see cref="PageNumberFinder"/> class.
        /// </summary>
        /// <param name="collector">A collector instance that has layout model records for the document.</param>
        public PageNumberFinder(LayoutCollector collector)
        {
            this.collector = collector;
        }

        /// <summary>
        /// Gets the document this instance works with.
        /// </summary>
        public Document Document => collector.Document;

        /// <summary>
        /// Retrieves 1-based index of a page that the node begins on.
        /// </summary>
        /// <param name="node">
        /// The node.
        /// </param>
        /// <returns>
        /// Page index.
        /// </returns>
        public int GetPage(Node node)
        {
            return nodeStartPageLookup.ContainsKey(node)
                ? nodeStartPageLookup[node]
                : collector.GetStartPageIndex(node);
        }

        /// <summary>
        /// Retrieves 1-based index of a page that the node ends on.
        /// </summary>
        /// <param name="node">
        /// The node.
        /// </param>
        /// <returns>
        /// Page index.
        /// </returns>
        public int GetPageEnd(Node node)
        {
            return nodeEndPageLookup.ContainsKey(node)
                ? nodeEndPageLookup[node]
                : collector.GetEndPageIndex(node);
        }

        /// <summary>
        /// Returns how many pages the specified node spans over. Returns 1 if the node is contained within one page.
        /// </summary>
        /// <param name="node">
        /// The node.
        /// </param>
        /// <returns>
        /// Page index.
        /// </returns>
        public int PageSpan(Node node)
        {
            return GetPageEnd(node) - GetPage(node) + 1;
        }

        /// <summary>
        /// Returns a list of nodes that are contained anywhere on the specified page or pages which match the specified node type.
        /// </summary>
        /// <param name="startPage">
        /// The start Page.
        /// </param>
        /// <param name="endPage">
        /// The end Page.
        /// </param>
        /// <param name="nodeType">
        /// The node Type.
        /// </param>
        /// <returns>
        /// The <see cref="IList{T}"/>.
        /// </returns>
        public IList<Node> RetrieveAllNodesOnPages(int startPage, int endPage, NodeType nodeType)
        {
            if (startPage < 1 || startPage > Document.PageCount)
            {
                throw new InvalidOperationException("'startPage' is out of range");
            }

            if (endPage < 1 || endPage > Document.PageCount || endPage < startPage)
            {
                throw new InvalidOperationException("'endPage' is out of range");
            }

            CheckPageListsPopulated();

            IList<Node> pageNodes = new List<Node>();
            for (int page = startPage; page <= endPage; page++)
            {
                // Some pages can be empty.
                if (!reversePageLookup.ContainsKey(page))
                {
                    continue;
                }

                foreach (Node node in reversePageLookup[page])
                {
                    if (node.ParentNode != null
                        && (nodeType == NodeType.Any || node.NodeType == nodeType)
                        && !pageNodes.Contains(node))
                    {
                        pageNodes.Add(node);
                    }
                }
            }

            return pageNodes;
        }

        /// <summary>
        /// Splits nodes that appear over two or more pages into separate nodes so that they still appear in the same way
        /// but no longer appear across a page.
        /// </summary>
        public void SplitNodesAcrossPages()
        {
            foreach (Paragraph paragraph in Document.GetChildNodes(NodeType.Paragraph, true))
            {
                if (GetPage(paragraph) != GetPageEnd(paragraph))
                {
                    SplitRunsByWords(paragraph);
                }
            }

            ClearCollector();

            // Visit any composites which are possibly split across pages and split them into separate nodes.
            Document.Accept(new SectionSplitter(this));
        }

        /// <summary>
        /// This is called by <see cref="SectionSplitter"/> to update page numbers of split nodes.
        /// </summary>
        /// <param name="node">
        /// The node.
        /// </param>
        /// <param name="startPage">
        /// The start Page.
        /// </param>
        /// <param name="endPage">
        /// The end Page.
        /// </param>
        internal void AddPageNumbersForNode(Node node, int startPage, int endPage)
        {
            if (startPage > 0)
            {
                nodeStartPageLookup[node] = startPage;
            }

            if (endPage > 0)
            {
                nodeEndPageLookup[node] = endPage;
            }
        }

        private bool IsHeaderFooterType(Node node)
        {
            return node.NodeType == NodeType.HeaderFooter || node.GetAncestor(NodeType.HeaderFooter) != null;
        }

        private void CheckPageListsPopulated()
        {
            if (reversePageLookup != null)
            {
                return;
            }

            reversePageLookup = new Dictionary<int, IList<Node>>();

            // Add each node to a list that represent the nodes found on each page.
            foreach (Node node in Document.GetChildNodes(NodeType.Any, true))
            {
                // Headers/Footers follow sections and are not split by themselves.
                if (IsHeaderFooterType(node))
                {
                    continue;
                }

                int startPage = GetPage(node);
                int endPage = GetPageEnd(node);
                for (int page = startPage; page <= endPage; page++)
                {
                    if (!reversePageLookup.ContainsKey(page))
                    {
                        reversePageLookup.Add(page, new List<Node>());
                    }

                    reversePageLookup[page].Add(node);
                }
            }
        }

        private void SplitRunsByWords(Paragraph paragraph)
        {
            foreach (Run run in paragraph.Runs)
            {
                if (GetPage(run) == GetPageEnd(run))
                {
                    continue;
                }

                SplitRunByWords(run);
            }
        }

        private void SplitRunByWords(Run run)
        {
            IEnumerable<string> words = run.Text.Split(' ').Reverse();

            foreach (string word in words)
            {
                int pos = run.Text.Length - word.Length - 1;
                if (pos > 1)
                {
                    SplitRun(run, run.Text.Length - word.Length - 1);
                }
            }
        }

        /// <summary>
        /// Splits text of the specified run into two runs.
        /// Inserts the new run just after the specified run.
        /// </summary>
        private void SplitRun(Run run, int position)
        {
            Run afterRun = (Run) run.Clone(true);
            afterRun.Text = run.Text.Substring(position);
            run.Text = run.Text.Substring(0, position);
            run.ParentNode.InsertAfter(afterRun, run);
        }

        private void ClearCollector()
        {
            collector.Clear();
            Document.UpdatePageLayout();

            nodeStartPageLookup.Clear();
            nodeEndPageLookup.Clear();
        }
    }

    internal class PageNumberFinderFactory
    {
        public static PageNumberFinder Create(Document document)
        {
            LayoutCollector layoutCollector = new LayoutCollector(document);
            document.UpdatePageLayout();
            PageNumberFinder pageNumberFinder = new PageNumberFinder(layoutCollector);
            pageNumberFinder.SplitNodesAcrossPages();
            return pageNumberFinder;
        }
    }

    /// <summary>
    /// Splits a document into multiple sections so that each page begins and ends at a section boundary.
    /// </summary>
    internal class SectionSplitter : DocumentVisitor
    {
        private readonly PageNumberFinder pageNumberFinder;

        public SectionSplitter(PageNumberFinder pageNumberFinder)
        {
            this.pageNumberFinder = pageNumberFinder;
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
            Section previousSection = (Section) section.PreviousSibling;

            // If there is a previous section, attempt to copy any linked header footers.
            // Otherwise, they will not appear in an extracted document if the previous section is missing.
            if (previousSection != null)
            {
                HeaderFooterCollection previousHeaderFooters = previousSection.HeadersFooters;
                if (!section.PageSetup.RestartPageNumbering)
                {
                    section.PageSetup.RestartPageNumbering = true;
                    section.PageSetup.PageStartingNumber = previousSection.PageSetup.PageStartingNumber +
                                                           pageNumberFinder.PageSpan(previousSection);
                }

                foreach (HeaderFooter previousHeaderFooter in previousHeaderFooters)
                {
                    if (section.HeadersFooters[previousHeaderFooter.HeaderFooterType] == null)
                    {
                        HeaderFooter newHeaderFooter =
                            (HeaderFooter) previousHeaderFooters[previousHeaderFooter.HeaderFooterType].Clone(true);
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
            // If the paragraph contains only section break, add fake run into.
            if (paragraph.IsEndOfSection && paragraph.ChildNodes.Count == 1 &&
                paragraph.ChildNodes[0].GetText() == "\f")
            {
                Run run = new Run(paragraph.Document);
                paragraph.AppendChild(run);
                int currentEndPageNum = pageNumberFinder.GetPageEnd(paragraph);
                pageNumberFinder.AddPageNumbersForNode(run, currentEndPageNum, currentEndPageNum);
            }

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
                cloneSection.PageSetup.PageStartingNumber = section.PageSetup.PageStartingNumber +
                                                            (section.Document.IndexOf(cloneSection) -
                                                             section.Document.IndexOf(section));
                cloneSection.PageSetup.DifferentFirstPageHeaderFooter = false;

                // Corrects page break at the end of the section.
                SplitPageBreakCorrector.ProcessSection(cloneSection);
            }

            SplitPageBreakCorrector.ProcessSection(section);

            // Add new page numbering for the body of the section as well.
            pageNumberFinder.AddPageNumbersForNode(section.Body, pageNumberFinder.GetPage(section),
                pageNumberFinder.GetPageEnd(section));
            return VisitorAction.Continue;
        }

        private VisitorAction ContinueIfCompositeAcrossPageElseSkip(CompositeNode composite)
        {
            return pageNumberFinder.PageSpan(composite) > 1
                ? VisitorAction.Continue
                : VisitorAction.SkipThisNode;
        }

        private List<Node> SplitComposite(CompositeNode composite)
        {
            List<Node> splitNodes = new List<Node>();
            foreach (Node splitNode in FindChildSplitPositions(composite))
            {
                splitNodes.Add(SplitCompositeAtNode(composite, splitNode));
            }

            return splitNodes;
        }

        private IEnumerable<Node> FindChildSplitPositions(CompositeNode node)
        {
            // A node may span across multiple pages, so a list of split positions is returned.
            // The split node is the first node on the next page.
            List<Node> splitList = new List<Node>();

            int startingPage = pageNumberFinder.GetPage(node);
            
            Node[] childNodes = node.NodeType == NodeType.Section
                ? ((Section) node).Body.ChildNodes.ToArray()
                : node.ChildNodes.ToArray();
            foreach (Node childNode in childNodes)
            {
                int pageNum = pageNumberFinder.GetPage(childNode);

                if (childNode is Run)
                {
                    pageNum = pageNumberFinder.GetPageEnd(childNode);
                }

                // If the page of the child node has changed, then this is the split position.
                // Add this to the list.
                if (pageNum > startingPage)
                {
                    splitList.Add(childNode);
                    startingPage = pageNum;
                }

                if (pageNumberFinder.PageSpan(childNode) > 1)
                {
                    pageNumberFinder.AddPageNumbersForNode(childNode, pageNum, pageNum);
                }
            }

            // Split composites backward, so the cloned nodes are inserted in the right order.
            splitList.Reverse();
            return splitList;
        }

        private CompositeNode SplitCompositeAtNode(CompositeNode baseNode, Node targetNode)
        {
            CompositeNode cloneNode = (CompositeNode) baseNode.Clone(false);
            Node node = targetNode;
            int currentPageNum = pageNumberFinder.GetPage(baseNode);

            // Move all nodes found on the next page into the copied node. Handle row nodes separately.
            if (baseNode.NodeType != NodeType.Row)
            {
                CompositeNode composite = cloneNode;
                if (baseNode.NodeType == NodeType.Section)
                {
                    cloneNode = (CompositeNode) baseNode.Clone(true);
                    Section section = (Section) cloneNode;
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
                // If we are dealing with a row, we need to add dummy cells for the cloned row.
                int targetPageNum = pageNumberFinder.GetPage(targetNode);
                
                Node[] childNodes = baseNode.ChildNodes.ToArray();
                foreach (Node childNode in childNodes)
                {
                    int pageNum = pageNumberFinder.GetPage(childNode);
                    if (pageNum == targetPageNum)
                    {
                        if (cloneNode.NodeType == NodeType.Row)
                            ((Row) cloneNode).EnsureMinimum();

                        if (cloneNode.NodeType == NodeType.Cell)
                            ((Cell) cloneNode).EnsureMinimum();

                        cloneNode.LastChild.Remove();
                        cloneNode.AppendChild(childNode);
                    }
                    else if (pageNum == currentPageNum)
                    {
                        cloneNode.AppendChild(childNode.Clone(false));
                        if (cloneNode.LastChild.NodeType != NodeType.Cell)
                        {
                            ((CompositeNode) cloneNode.LastChild).AppendChild(
                                ((CompositeNode) childNode).FirstChild.Clone(false));
                        }
                    }
                }
            }

            // Insert the split node after the original.
            baseNode.ParentNode.InsertAfter(cloneNode, baseNode);

            // Update the new page numbers of the base node and the cloned node, including its descendants.
            // This will only be a single page as the cloned composite is split to be on one page.
            int currentEndPageNum = pageNumberFinder.GetPageEnd(baseNode);
            pageNumberFinder.AddPageNumbersForNode(baseNode, currentPageNum, currentEndPageNum - 1);
            pageNumberFinder.AddPageNumbersForNode(cloneNode, currentEndPageNum, currentEndPageNum);
            foreach (Node childNode in cloneNode.GetChildNodes(NodeType.Any, true))
            {
                pageNumberFinder.AddPageNumbersForNode(childNode, currentEndPageNum, currentEndPageNum);
            }

            return cloneNode;
        }
    }

    internal class SplitPageBreakCorrector
    {
        private const string PageBreakStr = "\f";
        private const char PageBreak = '\f';

        public static void ProcessSection(Section section)
        {
            if (section.ChildNodes.Count == 0)
            {
                return;
            }

            Body lastBody = section.ChildNodes.OfType<Body>().LastOrDefault();

            Run run = lastBody?.GetChildNodes(NodeType.Run, true).OfType<Run>()
                .FirstOrDefault(p => p.Text.EndsWith(PageBreakStr));

            if (run != null)
            {
                RemovePageBreak(run);
            }
        }

        public void RemovePageBreakFromParagraph(Paragraph paragraph)
        {
            Run run = (Run) paragraph.FirstChild;
            if (run.Text == PageBreakStr)
            {
                paragraph.RemoveChild(run);
            }
        }

        private void ProcessLastParagraph(Paragraph paragraph)
        {
            Node lastNode = paragraph.ChildNodes[paragraph.ChildNodes.Count - 1];
            if (lastNode.NodeType != NodeType.Run)
            {
                return;
            }

            Run run = (Run) lastNode;
            RemovePageBreak(run);
        }

        private static void RemovePageBreak(Run run)
        {
            Paragraph paragraph = run.ParentParagraph;
            
            if (run.Text == PageBreakStr)
            {
                paragraph.RemoveChild(run);
            }
            else if (run.Text.EndsWith(PageBreakStr))
            {
                run.Text = run.Text.TrimEnd(PageBreak);
            }

            if (paragraph.ChildNodes.Count == 0)
            {
                CompositeNode parent = paragraph.ParentNode;
                parent.RemoveChild(paragraph);
            }
        }
    }
}