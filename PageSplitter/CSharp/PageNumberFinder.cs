//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.Collections;

using Aspose.Words;
using Aspose.Words.Layout;

namespace PageSplitter
{
    /// <summary>
    /// Provides methods for extracting nodes of a document which are rendered on a specified pages.
    /// </summary>
    public class PageNumberFinder
    {
        /// <summary>
        /// Initializes new instance of this class.
        /// </summary>
        /// <param name="collector">A collector instance which has layout model records for the document.</param>
        public PageNumberFinder(LayoutCollector collector)
        {
            mCollector = collector;
        }

        /// <summary>
        /// Retrieves 1-based index of a page that the node begins on.
        /// </summary>
        public int GetPage(Node node)
        {
            if (mNodeStartPageLookup.ContainsKey(node))
               return (int)mNodeStartPageLookup[node];

            return mCollector.GetStartPageIndex(node);
        }

        /// <summary>
        /// Retrieves 1-based index of a page that the node ends on.
        /// </summary>
        public int GetPageEnd(Node node)
        {
            if (mNodeEndPageLookup.ContainsKey(node))
                return (int)mNodeEndPageLookup[node];

            return mCollector.GetEndPageIndex(node);
        }

        /// <summary>
        /// Returns how many pages the specified node spans over. Returns 1 if the node is contained within one page.
        /// </summary>
        public int PageSpan(Node node)
        {
            return GetPageEnd(node) - GetPage(node) + 1;
        }

        /// <summary>
        /// Returns a list of nodes that are contained anywhere on the specified page or pages which match the specified node type.
        /// </summary>
        public ArrayList RetrieveAllNodesOnPages(int startPage, int endPage, NodeType nodeType)
        {
            if (startPage < 1 || startPage > Document.PageCount)
                throw new ArgumentOutOfRangeException("startPage");

            if (endPage < 1 || endPage > Document.PageCount || endPage < startPage)
                throw new ArgumentOutOfRangeException("endPage");

            CheckPageListsPopulated();

            ArrayList pageNodes = new ArrayList();

            for (int page = startPage; page <= endPage; page++)
            {
                // Some pages can be empty.
                if (!mReversePageLookup.ContainsKey(page))
                    continue;

                foreach (Node node in (ArrayList)mReversePageLookup[page])
                {
                    if (node.ParentNode != null && (nodeType == NodeType.Any || node.NodeType == nodeType) && !pageNodes.Contains(node))
                        pageNodes.Add(node);
                }
            }

            return pageNodes;
        }

        /// <summary>
        /// Splits nodes which appear over two or more pages into separate nodes so that they still appear in the same way
        /// but no longer appear across a page.
        /// </summary>
        public void SplitNodesAcrossPages()
        {
            // Visit any composites which are possibly split across pages and split them into separate nodes.
            Document.Accept(new SectionSplitter(this));
        }

        /// <summary>
        /// Gets the document this instance works with.
        /// </summary>
        public Document Document
        {
            get { return mCollector.Document; }
        }

        /// <summary>
        /// This is called by <see cref="SectionSplitter"/> to update page numbers of split nodes.
        /// </summary>
        internal void AddPageNumbersForNode(Node node, int startPage, int endPage)
        {
            if (startPage > 0)
                mNodeStartPageLookup[node] = startPage;

            if (endPage > 0)
                mNodeEndPageLookup[node] = endPage;
        }

        private void CheckPageListsPopulated()
        {
            if (mReversePageLookup != null)
                return;

            mReversePageLookup = new Hashtable();

            // Add each node to a list which represent the nodes found on each page.
            foreach (Node node in Document.GetChildNodes(NodeType.Any, true))
            {
                // Headers/Footers follow sections. They are not split by themselves.
                if (IsHeaderFooterType(node))
                    continue;

                int startPage = GetPage(node);
                int endPage = GetPageEnd(node);

                for (int page = startPage; page <= endPage; page++)
                {
                    if (!mReversePageLookup.ContainsKey(page))
                        mReversePageLookup.Add(page, new ArrayList());

                    ((ArrayList)mReversePageLookup[page]).Add(node);
                }
            }
        }

        private static bool IsHeaderFooterType(Node node)
        {
            return node.NodeType == NodeType.HeaderFooter || node.GetAncestor(NodeType.HeaderFooter) != null;
        }

        // Maps node to a start/end page numbers. This is used to override baseline page numbers provided by collector when document is split.
        private Hashtable mNodeStartPageLookup = new Hashtable();
        private Hashtable mNodeEndPageLookup = new Hashtable();
        // Maps page number to a list of nodes found on that page.
        private Hashtable mReversePageLookup;
        private LayoutCollector mCollector;
    }
}