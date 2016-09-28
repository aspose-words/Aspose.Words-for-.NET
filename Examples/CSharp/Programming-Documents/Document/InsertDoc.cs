using System.IO;
using Aspose.Words;
using System;
using Aspose.Words.MailMerging;
using System.Text.RegularExpressions;
using Aspose.Words.Replacing;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class InsertDoc
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            // Invokes the InsertDocument method shown above to insert a document at a bookmark.
            InsertDocumentAtBookmark(dataDir);
            InsertDocumentAtMailMerge(dataDir);
            InsertDocumentAtReplace(dataDir);
        }
        public static void InsertDocumentAtReplace(string dataDir)
        {
            //ExStart:InsertDocumentAtReplace
            Document mainDoc = new Document(dataDir + "InsertDocument1.doc");

            FindReplaceOptions options = new FindReplaceOptions();
            options.ReplacingCallback = new InsertDocumentAtReplaceHandler();

            mainDoc.Range.Replace(new Regex("\\[MY_DOCUMENT\\]"),"" , options);
            dataDir = dataDir + "InsertDocumentAtReplace_out_.doc";
            mainDoc.Save(dataDir);
            //ExEnd:InsertDocumentAtReplace
            Console.WriteLine("\nDocument inserted successfully at a replace.\nFile saved at " + dataDir);
        }
        public static void InsertDocumentAtBookmark(string dataDir)
        {
            //ExStart:InsertDocumentAtBookmark         
            Document mainDoc = new Document(dataDir + "InsertDocument1.doc");
            Document subDoc = new Document(dataDir + "InsertDocument2.doc");

            Bookmark bookmark = mainDoc.Range.Bookmarks["insertionPlace"];
            InsertDocument(bookmark.BookmarkStart.ParentNode, subDoc);
            dataDir = dataDir + "InsertDocumentAtBookmark_out_.doc";
            mainDoc.Save(dataDir);
            //ExEnd:InsertDocumentAtBookmark
            Console.WriteLine("\nDocument inserted successfully at a bookmark.\nFile saved at " + dataDir);
        }
        public static void InsertDocumentAtMailMerge(string dataDir)
        {
            //ExStart:InsertDocumentAtMailMerge   
            // Open the main document.
            Document mainDoc = new Document(dataDir + "InsertDocument1.doc");

            // Add a handler to MergeField event
            mainDoc.MailMerge.FieldMergingCallback = new InsertDocumentAtMailMergeHandler();

            // The main document has a merge field in it called "Document_1".
            // The corresponding data for this field contains fully qualified path to the document
            // that should be inserted to this field.
            mainDoc.MailMerge.Execute(
                new string[] { "Document_1" },
                new string[] { dataDir + "InsertDocument2.doc" });
            dataDir = dataDir + "InsertDocumentAtMailMerge_out_.doc";
            mainDoc.Save(dataDir);
            //ExEnd:InsertDocumentAtMailMerge 
            Console.WriteLine("\nDocument inserted successfully at mail merge.\nFile saved at " + dataDir);
        }
        //ExStart:InsertDocument
        /// <summary>
        /// Inserts content of the external document after the specified node.
        /// Section breaks and section formatting of the inserted document are ignored.
        /// </summary>
        /// <param name="insertAfterNode">Node in the destination document after which the content
        /// should be inserted. This node should be a block level node (paragraph or table).</param>
        /// <param name="srcDoc">The document to insert.</param>
        static void InsertDocument(Node insertAfterNode, Document srcDoc)
        {
            // Make sure that the node is either a paragraph or table.
            if ((!insertAfterNode.NodeType.Equals(NodeType.Paragraph)) &
              (!insertAfterNode.NodeType.Equals(NodeType.Table)))
                throw new ArgumentException("The destination node should be either a paragraph or table.");

            // We will be inserting into the parent of the destination paragraph.
            CompositeNode dstStory = insertAfterNode.ParentNode;

            // This object will be translating styles and lists during the import.
            NodeImporter importer = new NodeImporter(srcDoc, insertAfterNode.Document, ImportFormatMode.KeepSourceFormatting);

            // Loop through all sections in the source document.
            foreach (Section srcSection in srcDoc.Sections)
            {
                // Loop through all block level nodes (paragraphs and tables) in the body of the section.
                foreach (Node srcNode in srcSection.Body)
                {
                    // Let's skip the node if it is a last empty paragraph in a section.
                    if (srcNode.NodeType.Equals(NodeType.Paragraph))
                    {
                        Paragraph para = (Paragraph)srcNode;
                        if (para.IsEndOfSection && !para.HasChildNodes)
                            continue;
                    }

                    // This creates a clone of the node, suitable for insertion into the destination document.
                    Node newNode = importer.ImportNode(srcNode, true);

                    // Insert new node after the reference node.
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
        /// should be inserted. This node should be a block level node (paragraph or table).</param>
        /// <param name="srcDoc">The document to insert.</param>
        static void InsertDocumentWithSectionFormatting(Node insertAfterNode, Document srcDoc)
        {
            // Make sure that the node is either a pargraph or table.
            if ((!insertAfterNode.NodeType.Equals(NodeType.Paragraph)) &
                (!insertAfterNode.NodeType.Equals(NodeType.Table)))
                throw new ArgumentException("The destination node should be either a paragraph or table.");

            // Document to insert srcDoc into.
            Document dstDoc = (Document)insertAfterNode.Document;
            // To retain section formatting, split the current section into two at the marker node and then import the content from srcDoc as whole sections.
            // The section of the node which the insert marker node belongs to
            Section currentSection = (Section)insertAfterNode.GetAncestor(NodeType.Section);

            // Don't clone the content inside the section, we just want the properties of the section retained.
            Section cloneSection = (Section)currentSection.Clone(false);

            // However make sure the clone section has a body, but no empty first paragraph.
            cloneSection.EnsureMinimum();
            cloneSection.Body.FirstParagraph.Remove();

            // Insert the cloned section into the document after the original section.
            insertAfterNode.Document.InsertAfter(cloneSection, currentSection);

            // Append all nodes after the marker node to the new section. This will split the content at the section level at
            // the marker so the sections from the other document can be inserted directly.
            Node currentNode = insertAfterNode.NextSibling;
            while (currentNode != null)
            {
                Node nextNode = currentNode.NextSibling;
                cloneSection.Body.AppendChild(currentNode);
                currentNode = nextNode;
            }

            // This object will be translating styles and lists during the import.
            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.UseDestinationStyles);

            // Loop through all sections in the source document.
            foreach (Section srcSection in srcDoc.Sections)
            {
                Node newNode = importer.ImportNode(srcSection, true);

                // Append each section to the destination document. Start by inserting it after the split section.
                dstDoc.InsertAfter(newNode, currentSection);
                currentSection = (Section)newNode;
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
                    // Use document builder to navigate to the merge field with the specified name.
                    DocumentBuilder builder = new DocumentBuilder(e.Document);
                    builder.MoveToMergeField(e.DocumentFieldName);

                    // The name of the document to load and insert is stored in the field value.
                    Document subDoc = new Document((string)e.FieldValue);

                    // Insert the document.
                    InsertDocument(builder.CurrentParagraph, subDoc);

                    // The paragraph that contained the merge field might be empty now and you probably want to delete it.
                    if (!builder.CurrentParagraph.HasChildNodes)
                        builder.CurrentParagraph.Remove();

                    // Indicate to the mail merge engine that we have inserted what we wanted.
                    e.Text = null;
                }
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // Do nothing.
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
                    // Use document builder to navigate to the merge field with the specified name.
                    DocumentBuilder builder = new DocumentBuilder(e.Document);
                    builder.MoveToMergeField(e.DocumentFieldName);

                    // Load the document from the blob field.
                    MemoryStream stream = new MemoryStream((byte[])e.FieldValue);
                    Document subDoc = new Document(stream);

                    // Insert the document.
                    InsertDocument(builder.CurrentParagraph, subDoc);

                    // The paragraph that contained the merge field might be empty now and you probably want to delete it.
                    if (!builder.CurrentParagraph.HasChildNodes)
                        builder.CurrentParagraph.Remove();

                    // Indicate to the mail merge engine that we have inserted what we wanted.
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
                Document subDoc = new Document(RunExamples.GetDataDir_WorkingWithDocument() + "InsertDocument2.doc");

                // Insert a document after the paragraph, containing the match text.
                Paragraph para = (Paragraph)e.MatchNode.ParentNode;
                InsertDocument(para, subDoc);

                // Remove the paragraph with the match text.
                para.Remove();

                return ReplaceAction.Skip;
            }
        }
        //ExEnd:InsertDocumentAtReplaceHandler

    }

}
