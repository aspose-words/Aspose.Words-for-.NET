using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Replacing;
using NUnit.Framework;

namespace DocsExamples.Programming_with_Documents.Working_with_Document
{
    internal class CloneAndCombineDocuments : DocsExamplesBase
    {
        [Test]
        public void CloningDocument()
        {
            //ExStart:CloningDocument
            Document doc = new Document(MyDir + "Document.docx");

            Document clone = doc.Clone();
            clone.Save(ArtifactsDir + "CloneAndCombineDocuments.CloningDocument.docx");
            //ExEnd:CloningDocument
        }

        [Test]
        public void InsertDocumentAtReplace()
        {
            //ExStart:InsertDocumentAtReplace
            Document mainDoc = new Document(MyDir + "Document insertion 1.docx");

            FindReplaceOptions options = new FindReplaceOptions { ReplacingCallback = new InsertDocumentAtReplaceHandler() };

            mainDoc.Range.Replace(new Regex("\\[MY_DOCUMENT\\]"), "", options);

            mainDoc.Save(ArtifactsDir + "CloneAndCombineDocuments.InsertDocumentAtReplace.docx");
            //ExEnd:InsertDocumentAtReplace
        }

        [Test]
        public void InsertDocumentAtBookmark()
        {
            //ExStart:InsertDocumentAtBookmark         
            Document mainDoc = new Document(MyDir + "Document insertion 1.docx");
            Document subDoc = new Document(MyDir + "Document insertion 2.docx");

            Bookmark bookmark = mainDoc.Range.Bookmarks["insertionPlace"];
            InsertDocument(bookmark.BookmarkStart.ParentNode, subDoc);
            
            mainDoc.Save(ArtifactsDir + "CloneAndCombineDocuments.InsertDocumentAtBookmark.docx");
            //ExEnd:InsertDocumentAtBookmark
        }

        [Test]
        public void InsertDocumentAtMailMerge()
        {
            //ExStart:InsertDocumentAtMailMerge   
            Document mainDoc = new Document(MyDir + "Document insertion 1.docx");

            mainDoc.MailMerge.FieldMergingCallback = new InsertDocumentAtMailMergeHandler();
            // The main document has a merge field in it called "Document_1".
            // The corresponding data for this field contains a fully qualified path to the document.
            // That should be inserted to this field.
            mainDoc.MailMerge.Execute(new[] { "Document_1" }, new object[] { MyDir + "Document insertion 2.docx" });

            mainDoc.Save(ArtifactsDir + "CloneAndCombineDocuments.InsertDocumentAtMailMerge.doc");
            //ExEnd:InsertDocumentAtMailMerge
        }

        //ExStart:InsertDocument
        /// <summary>
        /// Inserts content of the external document after the specified node.
        /// Section breaks and section formatting of the inserted document are ignored.
        /// </summary>
        /// <param name="insertAfterNode">Node in the destination document after which the content
        /// Should be inserted. This node should be a block level node (paragraph or table).</param>
        /// <param name="srcDoc">The document to insert.</param>
        private static void InsertDocument(Node insertAfterNode, Document srcDoc)
        {
            if (!insertAfterNode.NodeType.Equals(NodeType.Paragraph) &
                !insertAfterNode.NodeType.Equals(NodeType.Table))
                throw new ArgumentException("The destination node should be either a paragraph or table.");

            CompositeNode dstStory = insertAfterNode.ParentNode;

            NodeImporter importer =
                new NodeImporter(srcDoc, insertAfterNode.Document, ImportFormatMode.KeepSourceFormatting);

            foreach (Section srcSection in srcDoc.Sections)
            {
                foreach (Node srcNode in srcSection.Body)
                {
                    // Let's skip the node if it is the last empty paragraph in a section.
                    if (srcNode.NodeType.Equals(NodeType.Paragraph))
                    {
                        Paragraph para = (Paragraph) srcNode;
                        if (para.IsEndOfSection && !para.HasChildNodes)
                            continue;
                    }

                    // This creates a clone of the node, suitable for insertion into the destination document.
                    Node newNode = importer.ImportNode(srcNode, true);

                    // Insert a new node after the reference node.
                    dstStory.InsertAfter(newNode, insertAfterNode);
                    insertAfterNode = newNode;
                }
            }
        }
        //ExEnd:InsertDocument

        //ExStart:InsertDocumentWithSectionFormatting
        /// <summary>
        /// Inserts content of the external document after the specified node.
        /// </summary>
        /// <param name="insertAfterNode">Node in the destination document after which the content
        /// Should be inserted. This node should be a block level node (paragraph or table).</param>
        /// <param name="srcDoc">The document to insert.</param>
        void InsertDocumentWithSectionFormatting(Node insertAfterNode, Document srcDoc)
        {
            if (!insertAfterNode.NodeType.Equals(NodeType.Paragraph) &
                !insertAfterNode.NodeType.Equals(NodeType.Table))
                throw new ArgumentException("The destination node should be either a paragraph or table.");

            Document dstDoc = (Document) insertAfterNode.Document;
            // To retain section formatting, split the current section into two at the marker node and then import the content
            // from srcDoc as whole sections. The section of the node to which the insert marker node belongs.
            Section currentSection = (Section) insertAfterNode.GetAncestor(NodeType.Section);

            // Don't clone the content inside the section, we just want the properties of the section retained.
            Section cloneSection = (Section) currentSection.Clone(false);

            // However, make sure the clone section has a body but no empty first paragraph.
            cloneSection.EnsureMinimum();
            cloneSection.Body.FirstParagraph.Remove();

            insertAfterNode.Document.InsertAfter(cloneSection, currentSection);

            // Append all nodes after the marker node to the new section. This will split the content at the section level at.
            // The marker so the sections from the other document can be inserted directly.
            Node currentNode = insertAfterNode.NextSibling;
            while (currentNode != null)
            {
                Node nextNode = currentNode.NextSibling;
                cloneSection.Body.AppendChild(currentNode);
                currentNode = nextNode;
            }

            // This object will be translating styles and lists during the import.
            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.UseDestinationStyles);

            foreach (Section srcSection in srcDoc.Sections)
            {
                Node newNode = importer.ImportNode(srcSection, true);

                dstDoc.InsertAfter(newNode, currentSection);
                currentSection = (Section) newNode;
            }
        }
        //ExEnd:InsertDocumentWithSectionFormatting

        //ExStart:InsertDocumentAtMailMergeHandler
        private class InsertDocumentAtMailMergeHandler : IFieldMergingCallback
        {
            /// <summary>
            /// This handler makes special processing for the "Document_1" field.
            /// The field value contains the path to load the document. 
            /// We load the document and insert it into the current merge field.
            /// </summary>
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
            {
                if (e.DocumentFieldName == "Document_1")
                {
                    DocumentBuilder builder = new DocumentBuilder(e.Document);
                    builder.MoveToMergeField(e.DocumentFieldName);

                    Document subDoc = new Document((string) e.FieldValue);
                    
                    InsertDocument(builder.CurrentParagraph, subDoc);

                    // The paragraph that contained the merge field might be empty now, and you probably want to delete it.
                    if (!builder.CurrentParagraph.HasChildNodes)
                        builder.CurrentParagraph.Remove();

                    e.Text = null;
                }
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // Do nothing
            }
        }
        //ExEnd:InsertDocumentAtMailMergeHandler

        //ExStart:InsertDocumentAtMailMergeBlobHandler
        private class InsertDocumentAtMailMergeBlobHandler : IFieldMergingCallback
        {
            /// <summary>
            /// This handler makes special processing for the "Document_1" field.
            /// The field value contains the path to load the document.
            /// We load the document and insert it into the current merge field.
            /// </summary>
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs e)
            {
                if (e.DocumentFieldName == "Document_1")
                {
                    DocumentBuilder builder = new DocumentBuilder(e.Document);
                    builder.MoveToMergeField(e.DocumentFieldName);

                    MemoryStream stream = new MemoryStream((byte[]) e.FieldValue);
                    Document subDoc = new Document(stream);

                    InsertDocument(builder.CurrentParagraph, subDoc);

                    // The paragraph that contained the merge field might be empty now, and you probably want to delete it.
                    if (!builder.CurrentParagraph.HasChildNodes)
                        builder.CurrentParagraph.Remove();

                    e.Text = null;
                }
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // Do nothing.
            }
        }
        //ExEnd:InsertDocumentAtMailMergeBlobHandler
        
        //ExStart:InsertDocumentAtReplaceHandler
        private class InsertDocumentAtReplaceHandler : IReplacingCallback
        {
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                Document subDoc = new Document(MyDir + "Document insertion 2.docx");

                Paragraph para = (Paragraph) e.MatchNode.ParentNode;
                InsertDocument(para, subDoc);

                para.Remove();

                return ReplaceAction.Skip;
            }
        }
        //ExEnd:InsertDocumentAtReplaceHandler
    }
}