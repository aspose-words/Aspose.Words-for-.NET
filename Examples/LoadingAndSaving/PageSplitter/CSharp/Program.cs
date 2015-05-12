//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;

using Aspose.Words;
using Aspose.Words.Layout;

namespace PageSplitterExample
{
    public class Program
    {
        public static void Main()
        {
            string dataDir = Path.GetFullPath("../../../Data/");
            SplitAllDocumentsToPages(dataDir);
        }

        public static void SplitDocumentToPages(string docName)
        {
            string folderName = Path.GetDirectoryName(docName);
            string fileName = Path.GetFileNameWithoutExtension(docName);
            string extensionName = Path.GetExtension(docName);
            string outFolder = Path.Combine(folderName, "Out");

            Console.WriteLine("Processing document: " + fileName + extensionName);

            Document doc = new Document(docName);

            // Create and attach collector to the document before page layout is built.
            LayoutCollector layoutCollector = new LayoutCollector(doc);

            // This will build layout model and collect necessary information.
            doc.UpdatePageLayout();

            // Split nodes in the document into separate pages.
            DocumentPageSplitter splitter = new DocumentPageSplitter(layoutCollector);

            // Save each page to the disk as a separate document.
            for (int page = 1; page <= doc.PageCount; page++)
            {
                Document pageDoc = splitter.GetDocumentOfPage(page);
                pageDoc.Save(Path.Combine(outFolder, string.Format("{0} - page{1} Out{2}", fileName, page, extensionName)));
            }

            // Detach the collector from the document.
            layoutCollector.Document = null;
        }

        public static void SplitAllDocumentsToPages(string folderName)
        {
            string[] fileNames = Directory.GetFiles(folderName, "*.doc?", SearchOption.TopDirectoryOnly);

            foreach (string fileName in fileNames)
            {
                SplitDocumentToPages(fileName);
            }
        }
    }
}