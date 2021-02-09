using System.IO;
using System.Linq;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Split_Documents
{
    internal class SplitDocument : DocsExamplesBase
    {
        [Test]
        public void ByHeadingsHtml()
        {
            //ExStart:SplitDocumentByHeadingsHtml
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions options = new HtmlSaveOptions
            {
                // Split a document into smaller parts, in this instance split by heading.
                DocumentSplitCriteria = DocumentSplitCriteria.HeadingParagraph
            };
            

            doc.Save(ArtifactsDir + "SplitDocument.ByHeadingsHtml.html", options);
            //ExEnd:SplitDocumentByHeadingsHtml
        }

        [Test]
        public void BySectionsHtml()
        {
            Document doc = new Document(MyDir + "Rendering.docx");
 
            //ExStart:SplitDocumentBySectionsHtml
            HtmlSaveOptions options = new HtmlSaveOptions { DocumentSplitCriteria = DocumentSplitCriteria.SectionBreak };
            //ExEnd:SplitDocumentBySectionsHtml
            
            doc.Save(ArtifactsDir + "SplitDocument.BySectionsHtml.html", options);
        }

        [Test]
        public void BySections()
        {
            //ExStart:SplitDocumentBySections
            Document doc = new Document(MyDir + "Big document.docx");

            for (int i = 0; i < doc.Sections.Count; i++)
            {
                // Split a document into smaller parts, in this instance, split by section.
                Section section = doc.Sections[i].Clone();

                Document newDoc = new Document();
                newDoc.Sections.Clear();

                Section newSection = (Section) newDoc.ImportNode(section, true);
                newDoc.Sections.Add(newSection);

                // Save each section as a separate document.
                newDoc.Save(ArtifactsDir + $"SplitDocument.BySections_{i}.docx");
            }
            //ExEnd:SplitDocumentBySections
        }

        [Test]
        public void PageByPage()
        {
            //ExStart:SplitDocumentPageByPage
            Document doc = new Document(MyDir + "Big document.docx");

            // Split nodes in the document into separate pages.
            DocumentPageSplitter splitter = new DocumentPageSplitter(doc);

            // Save each page as a separate document.
            for (int page = 1; page <= doc.PageCount; page++)
            {
                Document pageDoc = splitter.GetDocumentOfPage(page);
                pageDoc.Save(ArtifactsDir + $"SplitDocument.PageByPage_{page}.docx");
            }
            //ExEnd:SplitDocumentPageByPage

            MergeDocuments();
        }

        //ExStart:MergeSplitDocuments
        private void MergeDocuments()
        {
            // Find documents using for merge.
            FileSystemInfo[] documentPaths = new DirectoryInfo(ArtifactsDir)
                .GetFileSystemInfos("SplitDocument.PageByPage_*.docx").OrderBy(f => f.CreationTime).ToArray();
            string sourceDocumentPath =
                Directory.GetFiles(ArtifactsDir, "SplitDocument.PageByPage_1.docx", SearchOption.TopDirectoryOnly)[0];

            // Open the first part of the resulting document.
            Document sourceDoc = new Document(sourceDocumentPath);

            // Create a new resulting document.
            Document mergedDoc = new Document();
            DocumentBuilder mergedDocBuilder = new DocumentBuilder(mergedDoc);

            // Merge document parts one by one.
            foreach (FileSystemInfo documentPath in documentPaths)
            {
                if (documentPath.FullName == sourceDocumentPath)
                    continue;

                mergedDocBuilder.MoveToDocumentEnd();
                mergedDocBuilder.InsertDocument(sourceDoc, ImportFormatMode.KeepSourceFormatting);
                sourceDoc = new Document(documentPath.FullName);
            }

            mergedDoc.Save(ArtifactsDir + "SplitDocument.MergeDocuments.docx");
        }
        //ExEnd:MergeSplitDocuments

        [Test]
        public void ByPageRange()
        {
            //ExStart:SplitDocumentByPageRange
            Document doc = new Document(MyDir + "Big document.docx");
 
            // Split nodes in the document into separate pages.
            DocumentPageSplitter splitter = new DocumentPageSplitter(doc);
 
            // Get part of the document.
            Document pageDoc = splitter.GetDocumentOfPageRange(3,6);
            pageDoc.Save(ArtifactsDir + "SplitDocument.ByPageRange.docx");
            //ExEnd:SplitDocumentByPageRange
        }
    }
}

