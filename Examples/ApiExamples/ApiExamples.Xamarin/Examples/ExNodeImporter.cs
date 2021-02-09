// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
        [TestCase(false)]
        [TestCase(true)]
        public void KeepSourceNumbering(bool keepSourceNumbering)
        {
            //ExStart
            //ExFor:ImportFormatOptions.KeepSourceNumbering
            //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode, ImportFormatOptions)
            //ExSummary:Shows how to resolve list numbering clashes in source and destination documents.
            // Open a document with a custom list numbering scheme, and then clone it.
            // Since both have the same numbering format, the formats will clash if we import one document into the other.
            Document srcDoc = new Document(MyDir + "Custom list numbering.docx");
            Document dstDoc = srcDoc.Clone();

            // When we import the document's clone into the original and then append it,
            // then the two lists with the same list format will join.
            // If we set the "KeepSourceNumbering" flag to "false", then the list from the document clone
            // that we append to the original will carry on the numbering of the list we append it to.
            // This will effectively merge the two lists into one.
            // If we set the "KeepSourceNumbering" flag to "true", then the document clone
            // list will preserve its original numbering, making the two lists appear as separate lists. 
            ImportFormatOptions importFormatOptions = new ImportFormatOptions();
            importFormatOptions.KeepSourceNumbering = keepSourceNumbering;

            NodeImporter importer = new NodeImporter(srcDoc, dstDoc, ImportFormatMode.KeepDifferentStyles, importFormatOptions);
            foreach (Paragraph paragraph in srcDoc.FirstSection.Body.Paragraphs)
            {
                Node importedNode = importer.ImportNode(paragraph, true);
                dstDoc.FirstSection.Body.AppendChild(importedNode);
            }

            dstDoc.UpdateListLabels();

            if (keepSourceNumbering)
            {
                Assert.AreEqual(
                    "6. Item 1\r\n" +
                    "7. Item 2 \r\n" +
                    "8. Item 3\r\n" +
                    "9. Item 4\r\n" +
                    "6. Item 1\r\n" +
                    "7. Item 2 \r\n" +
                    "8. Item 3\r\n" +
                    "9. Item 4", dstDoc.FirstSection.Body.ToString(SaveFormat.Text).Trim());
            }
            else
            {
                Assert.AreEqual(
                    "6. Item 1\r\n" +
                    "7. Item 2 \r\n" +
                    "8. Item 3\r\n" +
                    "9. Item 4\r\n" +
                    "10. Item 1\r\n" +
                    "11. Item 2 \r\n" +
                    "12. Item 3\r\n" +
                    "13. Item 4", dstDoc.FirstSection.Body.ToString(SaveFormat.Text).Trim());
            }
            //ExEnd
        }

        //ExStart
        //ExFor:Paragraph.IsEndOfSection
        //ExFor:NodeImporter
        //ExFor:NodeImporter.#ctor(DocumentBase, DocumentBase, ImportFormatMode)
        //ExFor:NodeImporter.ImportNode(Node, Boolean)
        //ExSummary:Shows how to insert the contents of one document to a bookmark in another document.
        [Test]
        public void InsertAtBookmark()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.StartBookmark("InsertionPoint");
            builder.Write("We will insert a document here: ");
            builder.EndBookmark("InsertionPoint");

            Document docToInsert = new Document();
            builder = new DocumentBuilder(docToInsert);

            builder.Write("Hello world!");

            docToInsert.Save(ArtifactsDir + "NodeImporter.InsertAtMergeField.docx");

            Bookmark bookmark = doc.Range.Bookmarks["InsertionPoint"];
            InsertDocument(bookmark.BookmarkStart.ParentNode, docToInsert);

            Assert.AreEqual("We will insert a document here: " +
                            "\rHello world!", doc.GetText().Trim());
        }

        /// <summary>
        /// Inserts the contents of a document after the specified node.
        /// </summary>
        static void InsertDocument(Node insertionDestination, Document docToInsert)
        {
            if (insertionDestination.NodeType.Equals(NodeType.Paragraph) || insertionDestination.NodeType.Equals(NodeType.Table))
            {
                CompositeNode destinationParent = insertionDestination.ParentNode;

                NodeImporter importer =
                    new NodeImporter(docToInsert, insertionDestination.Document, ImportFormatMode.KeepSourceFormatting);

                // Loop through all block-level nodes in the section's body,
                // then clone and insert every node that is not the last empty paragraph of a section.
                foreach (Section srcSection in docToInsert.Sections.OfType<Section>())
                    foreach (Node srcNode in srcSection.Body)
                    {
                        if (srcNode.NodeType.Equals(NodeType.Paragraph))
                        {
                            Paragraph para = (Paragraph)srcNode;
                            if (para.IsEndOfSection && !para.HasChildNodes)
                                continue;
                        }

                        Node newNode = importer.ImportNode(srcNode, true);

                        destinationParent.InsertAfter(newNode, insertionDestination);
                        insertionDestination = newNode;
                    }
            }
            else
            {
                throw new ArgumentException("The destination node should be either a paragraph or table.");
            }
        }
        //ExEnd

        [Test]
        public void InsertAtMergeField()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("A document will appear here: ");
            builder.InsertField(" MERGEFIELD Document_1 ");

            Document subDoc = new Document();
            builder = new DocumentBuilder(subDoc);
            builder.Write("Hello world!");

            subDoc.Save(ArtifactsDir + "NodeImporter.InsertAtMergeField.docx");

            doc.MailMerge.FieldMergingCallback = new InsertDocumentAtMailMergeHandler();

            // The main document has a merge field in it called "Document_1".
            // Execute a mail merge using a data source that contains a local system filename
            // of the document that we wish to insert into the MERGEFIELD.
            doc.MailMerge.Execute(new string[] { "Document_1" },
                new object[] { ArtifactsDir + "NodeImporter.InsertAtMergeField.docx" });

            Assert.AreEqual("A document will appear here: \r" +
                            "Hello world!", doc.GetText().Trim());
        }

        /// <summary>
        /// If the mail merge encounters a MERGEFIELD with a specified name,
        /// this handler treats the current value of a mail merge data source as a local system filename of a document.
        /// The handler will insert the document in its entirety into the MERGEFIELD instead of the current merge value.
        /// </summary>
        private class InsertDocumentAtMailMergeHandler : IFieldMergingCallback
        {
            void IFieldMergingCallback.FieldMerging(FieldMergingArgs args)
            {
                if (args.DocumentFieldName == "Document_1")
                {
                    DocumentBuilder builder = new DocumentBuilder(args.Document);
                    builder.MoveToMergeField(args.DocumentFieldName);

                    Document subDoc = new Document((string)args.FieldValue);

                    InsertDocument(builder.CurrentParagraph, subDoc);

                    if (!builder.CurrentParagraph.HasChildNodes)
                        builder.CurrentParagraph.Remove();

                    args.Text = null;
                }
            }

            void IFieldMergingCallback.ImageFieldMerging(ImageFieldMergingArgs args)
            {
                // Do nothing.
            }
        }
    }
}