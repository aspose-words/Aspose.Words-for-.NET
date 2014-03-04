// Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
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
    /// Splits a document into multiple documents, one per page.
    /// </summary>
    public class DocumentPageSplitter
    {
        /// <summary>
        /// Initializes new instance of this class. This method splits the document into sections so that each page 
        /// begins and ends at a section boundary. It is recommended not to modify the document afterwards.
        /// </summary>
        /// <param name="collector">A collector instance which has layout model records for the document.</param>
        public DocumentPageSplitter(LayoutCollector collector)
        {
            mPageNumberFinder = new PageNumberFinder(collector);
            mPageNumberFinder.SplitNodesAcrossPages();
        }

        /// <summary>
        /// Gets the document of a page.
        /// </summary>
        /// <param name="pageIndex">1-based index of a page.</param>
        public Document GetDocumentOfPage(int pageIndex)
        {
            return GetDocumentOfPageRange(pageIndex, pageIndex);
        }

        /// <summary>
        /// Gets the document of a page range.
        /// </summary>
        /// <param name="startIndex">1-based index of the start page.</param>
        /// <param name="endIndex">1-based index of the end page.</param>
        public Document GetDocumentOfPageRange(int startIndex, int endIndex)
        {
            Document result = (Document)Document.Clone(false);

            foreach (Section section in mPageNumberFinder.RetrieveAllNodesOnPages(startIndex, endIndex, NodeType.Section))
                result.AppendChild(result.ImportNode(section, true));

            return result;
        }

        /// <summary>
        /// Gets the document this instance works with.
        /// </summary>
        private Document Document
        {
            get { return mPageNumberFinder.Document; }
        }

        private PageNumberFinder mPageNumberFinder;
    }
}
