using System;
using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Examples.CSharp.Loading_Saving;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class SplitDocument
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();

            SplitDocumentBySections(dataDir);
            SplitDocumentPageByPage(dataDir);
            SplitDocumentByPageRange(dataDir);
        }

        public static void SplitDocumentBySections(string dataDir)
        {
            //ExStart:SplitDocumentBySections
            // Open a Word document
            Document doc = new Document(dataDir + "TestFile (Split).docx");

            for (int i = 0; i < doc.Sections.Count; i++)
            {
                // Split a document into smaller parts, in this instance split by section
                Section section = doc.Sections[i].Clone();

                Document newDoc = new Document();
                newDoc.Sections.Clear();

                Section newSection = (Section) newDoc.ImportNode(section, true);
                newDoc.Sections.Add(newSection);

                // Save each section as a separate document
                newDoc.Save(dataDir + $"SplitDocumentBySectionsOut_{i}.docx");
            }
            //ExEnd:SplitDocumentBySections
        }

        public static void SplitDocumentPageByPage(string dataDir)
        {
            //ExStart:SplitDocumentPageByPage
            // Open a Word document
            Document doc = new Document(dataDir + "TestFile (Split).docx");

            // Split nodes in the document into separate pages
            DocumentPageSplitter splitter = new DocumentPageSplitter(doc);

            // Save each page as a separate document
            for (int page = 1; page <= doc.PageCount; page++)
            {
                Document pageDoc = splitter.GetDocumentOfPage(page);
                pageDoc.Save(dataDir + $"SplitDocumentPageByPageOut_{page}.docx");
            }
            //ExEnd:SplitDocumentPageByPage

            MergeDocuments(dataDir);
        }

        //ExStart:MergeSplitDocuments
        public static void MergeDocuments(string dataDir)
        {
            // Find documents using for merge
            FileSystemInfo[] documentPaths = new DirectoryInfo(dataDir)
                .GetFileSystemInfos("SplitDocumentPageByPageOut_*.docx").OrderBy(f => f.CreationTime).ToArray();
            string sourceDocumentPath =
                Directory.GetFiles(dataDir, "SplitDocumentPageByPageOut_1.docx", SearchOption.TopDirectoryOnly)[0];

            // Open the first part of the resulting document
            Document sourceDoc = new Document(sourceDocumentPath);

            // Create a new resulting document
            Document mergedDoc = new Document();
            DocumentBuilder mergedDocBuilder = new DocumentBuilder(mergedDoc);

            // Merge document parts one by one
            foreach (FileSystemInfo documentPath in documentPaths)
            {
                if (documentPath.FullName == sourceDocumentPath)
                    continue;

                mergedDocBuilder.MoveToDocumentEnd();
                mergedDocBuilder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);
                sourceDoc = new Document(documentPath.FullName);
            }

            // Save the output file
            mergedDoc.Save(dataDir + "MergeDocuments_out.docx");
        }
        //ExEnd:MergeSplitDocuments

        public static void SplitDocumentByPageRange(string dataDir)
        {
            //ExStart:SplitDocumentByPageRange
            // Open a Word document
            Document doc = new Document(dataDir + "TestFile (Split).docx");
 
            // Split nodes in the document into separate pages
            DocumentPageSplitter splitter = new DocumentPageSplitter(doc);
 
            // Get part of the document
            Document pageDoc = splitter.GetDocumentOfPageRange(3,6);
            pageDoc.Save(dataDir + "SplitDocumentByPageRangeOut.docx");
            //ExEnd:SplitDocumentByPageRange
        }
    }
}

