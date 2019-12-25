// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Linq;
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
        //ExFor:NodeImporter
        //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode)
        //ExFor:NodeImporter.ImportNode(Node, Boolean)
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
            // Make sure that the node is either a paragraph or table
            if ((!insertAfterNode.NodeType.Equals(NodeType.Paragraph)) &
                (!insertAfterNode.NodeType.Equals(NodeType.Table)))
                throw new ArgumentException("The destination node should be either a paragraph or table.");

            // We will be inserting into the parent of the destination paragraph
            CompositeNode dstStory = insertAfterNode.ParentNode;

            // This object will be translating styles and lists during the import
            NodeImporter importer =
                new NodeImporter(srcDoc, insertAfterNode.Document, ImportFormatMode.KeepSourceFormatting);

            // Loop through all sections in the source document
            foreach (Section srcSection in srcDoc.Sections.OfType<Section>())
            {
                // Loop through all block level nodes (paragraphs and tables) in the body of the section
                foreach (Node srcNode in srcSection.Body)
                {
                    // Let's skip the node if it is a last empty paragraph in a section
                    if (srcNode.NodeType.Equals(NodeType.Paragraph))
                    {
                        Paragraph para = (Paragraph) srcNode;
                        if (para.IsEndOfSection && !para.HasChildNodes)
                            continue;
                    }

                    // This creates a clone of the node, suitable for insertion into the destination document
                    Node newNode = importer.ImportNode(srcNode, true);

                    // Insert new node after the reference node
                    dstStory.InsertAfter(newNode, insertAfterNode);
                    insertAfterNode = newNode;
                }
            }
        }
        //ExEnd

        [Test]
        public void InsertAtBookmark()
        {
            Document mainDoc = new Document(MyDir + "InsertDocument1.doc");
            Document subDoc = new Document(MyDir + "InsertDocument2.doc");

            Bookmark bookmark = mainDoc.Range.Bookmarks["insertionPlace"];
            InsertDocument(bookmark.BookmarkStart.ParentNode, subDoc);

            mainDoc.Save(ArtifactsDir + "InsertDocument.InsertAtBookmark.doc");
        }

        //ExStart
        //ExFor:CompositeNode.HasChildNodes
        //ExSummary:Demonstrates how to use the InsertDocument method to insert a document into a merge field during mail merge.
        [Test] //ExSkip
        public void InsertAtMailMerge()
        {
            // Open the main document
            Document mainDoc = new Document(MyDir + "InsertDocument1.doc");

            // Add a handler to MergeField event
            mainDoc.MailMerge.FieldMergingCallback = new InsertDocumentAtMailMergeHandler();

            // The main document has a merge field in it called "Document_1"
            // The corresponding data for this field contains fully qualified path to the document
            // that should be inserted to this field
            mainDoc.MailMerge.Execute(new string[] { "Document_1" }, new object[] { MyDir + "InsertDocument2.doc" });

            mainDoc.Save(ArtifactsDir + "InsertDocument.InsertAtMailMerge.doc");
        }

        private class InsertDocumentAtMailMergeHandler : IFieldMergingCallback
        {
            /// <summary>
            /// This handler makes special processing for the "Document_1" field.
            /// The field value contains the path to load the document. 
            /// We load the document and insert it into the current merge field.
            /// </summary>
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
            {
                if (args.DocumentFieldName == "Document_1")
                {
                    // Use document builder to navigate to the merge field with the specified name
                    DocumentBuilder builder = new DocumentBuilder(args.Document);
                    builder.MoveToMergeField(args.DocumentFieldName);

                    // The name of the document to load and insert is stored in the field value
                    Document subDoc = new Document((string) args.FieldValue);

                    // Insert the document
                    InsertDocument(builder.CurrentParagraph, subDoc);

                    // The paragraph that contained the merge field might be empty now and you probably want to delete it
                    if (!builder.CurrentParagraph.HasChildNodes)
                        builder.CurrentParagraph.Remove();

                    // Indicate to the mail merge engine that we have inserted what we wanted
                    args.Text = null;
                }
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // Do nothing
            }
        }
        //ExEnd
        
        //ExStart
        //ExFor:Range.Replace(Regex, String, FindReplaceOptions)
        //ExFor:IReplacingCallback
        //ExFor:ReplaceAction
        //ExFor:IReplacingCallback.Replacing
        //ExFor:ReplacingArgs
        //ExFor:ReplacingArgs.MatchNode
        //ExFor:FindReplaceDirection
        //ExSummary:Shows how to insert content of one document into another during a customized find and replace operation.
        [Test] //ExSkip
        public void InsertDocumentAtReplace()
        {
            Document mainDoc = new Document(MyDir + "InsertDocument1.doc");

            FindReplaceOptions options = new FindReplaceOptions();
            options.Direction = FindReplaceDirection.Backward;
            options.ReplacingCallback = new InsertDocumentAtReplaceHandler();

            mainDoc.Range.Replace(new Regex("\\[MY_DOCUMENT\\]"), "", options);
            mainDoc.Save(ArtifactsDir + "InsertDocument.InsertDocumentAtReplace.doc");
        }

        private class InsertDocumentAtReplaceHandler : IReplacingCallback
        {
            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
            {
                Document subDoc = new Document(MyDir + "InsertDocument2.doc");

                // Insert a document after the paragraph, containing the match text
                Paragraph para = (Paragraph) args.MatchNode.ParentNode;
                InsertDocument(para, subDoc);

                // Remove the paragraph with the match text
                para.Remove();

                return ReplaceAction.Skip;
            }
        }
        //ExEnd
    }
}