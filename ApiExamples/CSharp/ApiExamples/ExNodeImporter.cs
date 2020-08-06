// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Linq;
using Aspose.Words;
using Aspose.Words.MailMerging;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExNodeImporter : ApiExampleBase
    {
        //ExStart
        //ExFor:Paragraph.IsEndOfSection
        //ExFor:NodeImporter
        //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode)
        //ExFor:NodeImporter.ImportNode(Node, Boolean)
        //ExSummary:Shows how to insert the contents of one document to a bookmark in another document.
        [Test]
        public void InsertAtBookmark()
        {
            Document mainDoc = new Document(MyDir + "Document insertion destination.docx");
            Document docToInsert = new Document(MyDir + "Document.docx");

            Bookmark bookmark = mainDoc.Range.Bookmarks["insertionPlace"];
            InsertDocument(bookmark.BookmarkStart.ParentNode, docToInsert);

            mainDoc.Save(ArtifactsDir + "NodeImporter.InsertAtBookmark.docx");
            TestInsertAtBookmark(new Document(ArtifactsDir + "NodeImporter.InsertAtBookmark.docx")); //ExSkip
        }

        [Test]
        public void KeepSourceNumbering()
        {
            //ExStart
            //ExFor:ImportFormatOptions.KeepSourceNumbering
            //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode, ImportFormatOptions)
            //ExSummary:Shows how the numbering will be imported when it clashes in source and destination documents.
            // Open a document with a custom list numbering scheme and clone it
            // Since both have the same numbering format, the formats will clash if we import one document into the other
            Document srcDoc = new Document(MyDir + "Custom list numbering.docx");
            Document dstDoc = srcDoc.Clone();

            // Both documents have the same numbering in their lists, but if we set this flag to false and then import one document into the other
            // the numbering of the imported source document will continue from where it ends in the destination document
            ImportFormatOptions importFormatOptions = new ImportFormatOptions();
            importFormatOptions.KeepSourceNumbering = false;

            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepDifferentStyles, importFormatOptions);
            foreach (Paragraph paragraph in srcDoc.FirstSection.Body.Paragraphs)
            {
                Node importedNode = importer.ImportNode(paragraph, true);
                dstDoc.FirstSection.Body.AppendChild(importedNode);
            }

            dstDoc.UpdateListLabels();
            dstDoc.Save(ArtifactsDir + "NodeImporter.KeepSourceNumbering.docx");
            //ExEnd
        }

        /// <summary>
        /// Inserts content of the external document after the specified node.
        /// </summary>
        static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            // Make sure that the node is either a paragraph or table
            if (insertionDestination.NodeType.Equals(NodeType.Paragraph) || insertionDestination.NodeType.Equals(NodeType.Table))
            {
                // We will be inserting into the parent of the destination paragraph
                CompositeNode dstStory = insertionDestination.ParentNode;

                // This object will be translating styles and lists during the import
                NodeImporter importer =
                    new NodeImporter(docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

                // Loop through all block level nodes in the body of the section
                foreach (Section srcSection in docToInsert.Sections.OfType<Section>())
                    foreach (Node srcNode in srcSection.Body)
                    {
                        // Skip the node if it is a last empty paragraph in a section
                        if (srcNode.NodeType.Equals(NodeType.Paragraph))
                        {
                            Paragraph para = (Paragraph)srcNode;
                            if (para.IsEndOfSection && !para.HasChildNodes)
                                continue;
                        }

                        // This creates a clone of the node, suitable for insertion into the destination document
                        Node newNode = importer.ImportNode(srcNode, true);

                        // Insert new node after the reference node
                        dstStory.InsertAfter(newNode, insertionDestination);
                        insertionDestination = newNode;
                    }
            }
            else
            {
                throw new ArgumentException("The destination node should be either a paragraph or table.");
            }
        }
        //ExEnd

        private void TestInsertAtBookmark(Document doc)
        {
            Assert.AreEqual("1) At text that can be identified by regex:\r[MY_DOCUMENT]\r" +
                            "2) At a MERGEFIELD:\r\u0013 MERGEFIELD  Document_1  \\* MERGEFORMAT \u0014«Document_1»\u0015\r" +
                            "3) At a bookmark:\r\rHello World!", doc.FirstSection.Body.GetText().Trim());
        }

        [Test]
        public void InsertAtMailMerge()
        {
            // Open the main document
            Document mainDoc = new Document(MyDir + "Document insertion destination.docx");

            // Add a handler to MergeField event
            mainDoc.MailMerge.FieldMergingCallback = new InsertDocumentAtMailMergeHandler();

            // The main document has a merge field in it called "Document_1"
            // The corresponding data for this field contains fully qualified path to the document
            // that should be inserted to this field
            mainDoc.MailMerge.Execute(new string[] { "Document_1" }, new object[] { MyDir + "Document.docx" });

            mainDoc.Save(ArtifactsDir + "NodeImporter.InsertAtMailMerge.docx");
            TestInsertAtMailMerge(new Document(ArtifactsDir + "NodeImporter.InsertAtMailMerge.docx")); //ExSkip
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
                    Document subDoc = new Document((string)args.FieldValue);

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

        private void TestInsertAtMailMerge(Document doc)
        {
            Assert.AreEqual("1) At text that can be identified by regex:\r[MY_DOCUMENT]\r" +
                            "2) At a MERGEFIELD:\rHello World!\r" +
                            "3) At a bookmark:", doc.FirstSection.Body.GetText().Trim());
        }
    }
}