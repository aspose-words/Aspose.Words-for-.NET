// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using System.Text.RegularExpressions;

using Aspose.Words;
using Aspose.Words.MailMerging;
using Aspose.Words.Replacing;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExInsertDocument : ApiExampleBase
    {
        //ExStart
        //ExFor:Paragraph.IsEndOfSection
        //ExId:InsertDocumentMain
        //ExSummary:This is a method that inserts contents of one document at a specified location in another document.
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
        //ExEnd

        [Test]
        public void InsertDocumentAtBookmark()
        {
            //ExStart
            //ExId:InsertDocumentAtBookmark
            //ExSummary:Invokes the InsertDocument method shown above to insert a document at a bookmark.
            Document mainDoc = new Document(MyDir + "InsertDocument1.doc");
            Document subDoc = new Document(MyDir + "InsertDocument2.doc");

            Bookmark bookmark = mainDoc.Range.Bookmarks["insertionPlace"];
            InsertDocument(bookmark.BookmarkStart.ParentNode, subDoc);

            mainDoc.Save(MyDir + @"\Artifacts\InsertDocumentAtBookmark.doc");
            //ExEnd
        }

        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void InsertDocumentAtMailMergeCaller()
        {
            this.InsertDocumentAtMailMerge();
        }

        //ExStart
        //ExFor:CompositeNode.HasChildNodes
        //ExId:InsertDocumentAtMailMerge
        //ExSummary:Demonstrates how to use the InsertDocument method to insert a document into a merge field during mail merge.
        public void InsertDocumentAtMailMerge()
        {
            // Open the main document.
            Document mainDoc = new Document(MyDir + "InsertDocument1.doc");

            // Add a handler to MergeField event
            mainDoc.MailMerge.FieldMergingCallback = new InsertDocumentAtMailMergeHandler();

            // The main document has a merge field in it called "Document_1".
            // The corresponding data for this field contains fully qualified path to the document
            // that should be inserted to this field.
            mainDoc.MailMerge.Execute(
                new string[] { "Document_1" },
                new string[] { MyDir + "InsertDocument2.doc" });

            mainDoc.Save(MyDir + @"\Artifacts\InsertDocumentAtMailMerge.doc");
        }

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
        //ExEnd

        //ExStart
        //ExId:InsertDocumentAtMailMergeBlob
        //ExSummary:A slight variation to the above example to load a document from a BLOB database field instead of a file.
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
        //ExEnd

        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void InsertDocumentAtReplaceCaller()
        {
            this.InsertDocumentAtReplace();
        }
        
        //ExStart
        //ExFor:Range.Replace(Regex,IReplacingCallback,Boolean)
        //ExFor:IReplacingCallback
        //ExFor:ReplaceAction
        //ExFor:IReplacingCallback.Replacing
        //ExFor:ReplacingArgs
        //ExFor:ReplacingArgs.MatchNode
        //ExId:InsertDocumentAtReplace
        //ExSummary:Shows how to insert content of one document into another during a customized find and replace operation.
        public void InsertDocumentAtReplace()
        {
            Document mainDoc = new Document(MyDir + "InsertDocument1.doc");
            mainDoc.Range.Replace(new Regex("\\[MY_DOCUMENT\\]"), new InsertDocumentAtReplaceHandler(), false);
            mainDoc.Save(MyDir + @"\Artifacts\InsertDocumentAtReplace.doc");
        }

        private class InsertDocumentAtReplaceHandler : IReplacingCallback
        {
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
            {
                Document subDoc = new Document(MyDir + "InsertDocument2.doc");

                // Insert a document after the paragraph, containing the match text.
                Paragraph para = (Paragraph)e.MatchNode.ParentNode;
                InsertDocument(para, subDoc);

                // Remove the paragraph with the match text.
                para.Remove();

                return ReplaceAction.Skip;
            }
        }
        //ExEnd
    }
}
