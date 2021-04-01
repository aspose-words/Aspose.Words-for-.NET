// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Markup;
using Aspose.Words.Math;
using Aspose.Words.Notes;
using Aspose.Words.Tables;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExDocumentVisitor : ApiExampleBase
    {
        //ExStart
        //ExFor:Document.Accept(DocumentVisitor)
        //ExFor:Body.Accept(DocumentVisitor)
        //ExFor:SubDocument.Accept(DocumentVisitor)
        //ExFor:DocumentVisitor
        //ExFor:DocumentVisitor.VisitRun(Run)
        //ExFor:DocumentVisitor.VisitDocumentEnd(Document)
        //ExFor:DocumentVisitor.VisitDocumentStart(Document)
        //ExFor:DocumentVisitor.VisitSectionEnd(Section)
        //ExFor:DocumentVisitor.VisitSectionStart(Section)
        //ExFor:DocumentVisitor.VisitBodyStart(Body)
        //ExFor:DocumentVisitor.VisitBodyEnd(Body)
        //ExFor:DocumentVisitor.VisitParagraphStart(Paragraph)
        //ExFor:DocumentVisitor.VisitParagraphEnd(Paragraph)
        //ExFor:DocumentVisitor.VisitSubDocument(SubDocument)
        //ExSummary:Shows how to use a document visitor to print a document's node structure.
        [Test] //ExSkip
        public void DocStructureToText()
        {
            Document doc = new Document(MyDir + "DocumentVisitor-compatible features.docx");
            DocStructurePrinter visitor = new DocStructurePrinter();

            // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
            // and then traverses all the node's children in a depth-first manner.
            // The visitor can read and modify each visited node.
            doc.Accept(visitor);

            Console.WriteLine(visitor.GetText());
            TestDocStructureToText(visitor); //ExSkip
        }

        /// <summary>
        /// Traverses a node's tree of child nodes.
        /// Creates a map of this tree in the form of a string.
        /// </summary>
        public class DocStructurePrinter : DocumentVisitor
        {
            public DocStructurePrinter()
            {
                mAcceptingNodeChildTree = new StringBuilder();
            }

            public string GetText()
            {
                return mAcceptingNodeChildTree.ToString();
            }

            /// <summary>
            /// Called when a Document node is encountered.
            /// </summary>
            public override VisitorAction VisitDocumentStart(Document doc)
            {
                int childNodeCount = doc.GetChildNodes(NodeType.Any, true).Count;

                IndentAndAppendLine("[Document start] Child nodes: " + childNodeCount);
                mDocTraversalDepth++;

                // Allow the visitor to continue visiting other nodes.
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called after all the child nodes of a Document node have been visited.
            /// </summary>
            public override VisitorAction VisitDocumentEnd(Document doc)
            {
                mDocTraversalDepth--;
                IndentAndAppendLine("[Document end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Section node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitSectionStart(Section section)
            {
                // Get the index of our section within the document.
                NodeCollection docSections = section.Document.GetChildNodes(NodeType.Section, false);
                int sectionIndex = docSections.IndexOf(section);

                IndentAndAppendLine("[Section start] Section index: " + sectionIndex);
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called after all the child nodes of a Section node have been visited.
            /// </summary>
            public override VisitorAction VisitSectionEnd(Section section)
            {
                mDocTraversalDepth--;
                IndentAndAppendLine("[Section end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Body node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitBodyStart(Body body)
            {
                int paragraphCount = body.Paragraphs.Count;
                IndentAndAppendLine("[Body start] Paragraphs: " + paragraphCount);
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called after all the child nodes of a Body node have been visited.
            /// </summary>
            public override VisitorAction VisitBodyEnd(Body body)
            {
                mDocTraversalDepth--;
                IndentAndAppendLine("[Body end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Paragraph node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitParagraphStart(Paragraph paragraph)
            {
                IndentAndAppendLine("[Paragraph start]");
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called after all the child nodes of a Paragraph node have been visited.
            /// </summary>
            public override VisitorAction VisitParagraphEnd(Paragraph paragraph)
            {
                mDocTraversalDepth--;
                IndentAndAppendLine("[Paragraph end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Run node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitRun(Run run)
            {
                IndentAndAppendLine("[Run] \"" + run.GetText() + "\"");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a SubDocument node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitSubDocument(SubDocument subDocument)
            {
                IndentAndAppendLine("[SubDocument]");
                
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
            /// </summary>
            /// <param name="text"></param>
            private void IndentAndAppendLine(string text)
            {
                for (int i = 0; i < mDocTraversalDepth; i++) mAcceptingNodeChildTree.Append("|  ");

                mAcceptingNodeChildTree.AppendLine(text);
            }

            private int mDocTraversalDepth;
            private readonly StringBuilder mAcceptingNodeChildTree;
        }
        //ExEnd

        private static void TestDocStructureToText(DocStructurePrinter visitor)
        {
            string visitorText = visitor.GetText();

            Assert.True(visitorText.Contains("[Document start]"));
            Assert.True(visitorText.Contains("[Document end]"));
            Assert.True(visitorText.Contains("[Section start]"));
            Assert.True(visitorText.Contains("[Section end]"));
            Assert.True(visitorText.Contains("[Body start]"));
            Assert.True(visitorText.Contains("[Body end]"));
            Assert.True(visitorText.Contains("[Paragraph start]"));
            Assert.True(visitorText.Contains("[Paragraph end]"));
            Assert.True(visitorText.Contains("[Run]"));
            Assert.True(visitorText.Contains("[SubDocument]"));
        }

        //ExStart
        //ExFor:Cell.Accept(DocumentVisitor)
        //ExFor:Cell.IsFirstCell
        //ExFor:Cell.IsLastCell
        //ExFor:DocumentVisitor.VisitTableEnd(Tables.Table)
        //ExFor:DocumentVisitor.VisitTableStart(Tables.Table)
        //ExFor:DocumentVisitor.VisitRowEnd(Tables.Row)
        //ExFor:DocumentVisitor.VisitRowStart(Tables.Row)
        //ExFor:DocumentVisitor.VisitCellStart(Tables.Cell)
        //ExFor:DocumentVisitor.VisitCellEnd(Tables.Cell)
        //ExFor:Row.Accept(DocumentVisitor)
        //ExFor:Row.FirstCell
        //ExFor:Row.GetText
        //ExFor:Row.IsFirstRow
        //ExFor:Row.LastCell
        //ExFor:Row.ParentTable
        //ExSummary:Shows how to print the node structure of every table in a document.
        [Test] //ExSkip
        public void TableToText()
        {
            Document doc = new Document(MyDir + "DocumentVisitor-compatible features.docx");
            TableStructurePrinter visitor = new TableStructurePrinter();

            // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
            // and then traverses all the node's children in a depth-first manner.
            // The visitor can read and modify each visited node.
            doc.Accept(visitor);

            Console.WriteLine(visitor.GetText());
            TestTableToText(visitor); //ExSkip
        }

        /// <summary>
        /// Traverses a node's non-binary tree of child nodes.
        /// Creates a map in the form of a string of all encountered Table nodes and their children.
        /// </summary>
        public class TableStructurePrinter : DocumentVisitor
        {
            public TableStructurePrinter()
            {
                mVisitedTables = new StringBuilder();
                mVisitorIsInsideTable = false;
            }

            public string GetText()
            {
                return mVisitedTables.ToString();
            }

            /// <summary>
            /// Called when a Run node is encountered in the document.
            /// Runs that are not within tables are not recorded.
            /// </summary>
            public override VisitorAction VisitRun(Run run)
            {
                if (mVisitorIsInsideTable) IndentAndAppendLine("[Run] \"" + run.GetText() + "\"");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Table is encountered in the document.
            /// </summary>
            public override VisitorAction VisitTableStart(Table table)
            {
                int rows = 0;
                int columns = 0;

                if (table.Rows.Count > 0)
                {
                    rows = table.Rows.Count;
                    columns = table.FirstRow.Count;
                }

                IndentAndAppendLine("[Table start] Size: " + rows + "x" + columns);
                mDocTraversalDepth++;
                mVisitorIsInsideTable = true;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called after all the child nodes of a Table node have been visited.
            /// </summary>
            public override VisitorAction VisitTableEnd(Table table)
            {
                mDocTraversalDepth--;
                IndentAndAppendLine("[Table end]");
                mVisitorIsInsideTable = false;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Row node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitRowStart(Row row)
            {
                string rowContents = row.GetText().TrimEnd(new []{ '\u0007', ' ' }).Replace("\u0007", ", ");
                int rowWidth = row.IndexOf(row.LastCell) + 1;
                int rowIndex = row.ParentTable.IndexOf(row);
                string rowStatusInTable = row.IsFirstRow && row.IsLastRow ? "only" : row.IsFirstRow ? "first" : row.IsLastRow ? "last" : "";
                if (rowStatusInTable != "")
                {
                    rowStatusInTable = $", the {rowStatusInTable} row in this table,";
                }

                IndentAndAppendLine($"[Row start] Row #{++rowIndex}{rowStatusInTable} width {rowWidth}, \"{rowContents}\"");
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called after all the child nodes of a Row node have been visited.
            /// </summary>
            public override VisitorAction VisitRowEnd(Row row)
            {
                mDocTraversalDepth--;
                IndentAndAppendLine("[Row end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Cell node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitCellStart(Cell cell)
            {
                Row row = cell.ParentRow;
                Table table = row.ParentTable;
                string cellStatusInRow = cell.IsFirstCell && cell.IsLastCell ? "only" : cell.IsFirstCell ? "first" : cell.IsLastCell ? "last" : "";
                if (cellStatusInRow != "")
                {
                    cellStatusInRow = $", the {cellStatusInRow} cell in this row";
                }

                IndentAndAppendLine($"[Cell start] Row {table.IndexOf(row) + 1}, Col {row.IndexOf(cell) + 1}{cellStatusInRow}");
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called after all the child nodes of a Cell node have been visited.
            /// </summary>
            public override VisitorAction VisitCellEnd(Cell cell)
            {
                mDocTraversalDepth--;
                IndentAndAppendLine("[Cell end]");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Append a line to the StringBuilder, and indent it depending on how deep the visitor is
            /// into the current table's tree of child nodes.
            /// </summary>
            /// <param name="text"></param>
            private void IndentAndAppendLine(string text)
            {
                for (int i = 0; i < mDocTraversalDepth; i++)
                {
                    mVisitedTables.Append("|  ");
                }

                mVisitedTables.AppendLine(text);
            }

            private bool mVisitorIsInsideTable;
            private int mDocTraversalDepth;
            private readonly StringBuilder mVisitedTables;
        }
        //ExEnd

        private static void TestTableToText(TableStructurePrinter visitor)
        {
            string visitorText = visitor.GetText();

            Assert.True(visitorText.Contains("[Table start]"));
            Assert.True(visitorText.Contains("[Table end]"));
            Assert.True(visitorText.Contains("[Row start]"));
            Assert.True(visitorText.Contains("[Row end]"));
            Assert.True(visitorText.Contains("[Cell start]"));
            Assert.True(visitorText.Contains("[Cell end]"));
            Assert.True(visitorText.Contains("[Run]"));
        }

        //ExStart
        //ExFor:DocumentVisitor.VisitCommentStart(Comment)
        //ExFor:DocumentVisitor.VisitCommentEnd(Comment)
        //ExFor:DocumentVisitor.VisitCommentRangeEnd(CommentRangeEnd)
        //ExFor:DocumentVisitor.VisitCommentRangeStart(CommentRangeStart)
        //ExSummary:Shows how to print the node structure of every comment and comment range in a document.
        [Test] //ExSkip
        public void CommentsToText()
        {
            Document doc = new Document(MyDir + "DocumentVisitor-compatible features.docx");
            CommentStructurePrinter visitor = new CommentStructurePrinter();

            // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
            // and then traverses all the node's children in a depth-first manner.
            // The visitor can read and modify each visited node.
            doc.Accept(visitor);

            Console.WriteLine(visitor.GetText());
            TestCommentsToText(visitor); //ExSkip
        }

        /// <summary>
        /// Traverses a node's non-binary tree of child nodes.
        /// Creates a map in the form of a string of all encountered Comment/CommentRange nodes and their children.
        /// </summary>
        public class CommentStructurePrinter : DocumentVisitor
        {
            public CommentStructurePrinter()
            {
                mBuilder = new StringBuilder();
                mVisitorIsInsideComment = false;
            }

            public string GetText()
            {
                return mBuilder.ToString();
            }

            /// <summary>
            /// Called when a Run node is encountered in the document.
            /// A Run is only recorded if it is a child of a Comment or CommentRange node.
            /// </summary>
            public override VisitorAction VisitRun(Run run)
            {
                if (mVisitorIsInsideComment) IndentAndAppendLine("[Run] \"" + run.GetText() + "\"");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a CommentRangeStart node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitCommentRangeStart(CommentRangeStart commentRangeStart)
            {
                IndentAndAppendLine("[Comment range start] ID: " + commentRangeStart.Id);
                mDocTraversalDepth++;
                mVisitorIsInsideComment = true;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a CommentRangeEnd node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitCommentRangeEnd(CommentRangeEnd commentRangeEnd)
            {
                mDocTraversalDepth--;
                IndentAndAppendLine("[Comment range end]");
                mVisitorIsInsideComment = false;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Comment node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitCommentStart(Comment comment)
            {
                IndentAndAppendLine(
                    $"[Comment start] For comment range ID {comment.Id}, By {comment.Author} on {comment.DateTime}");
                mDocTraversalDepth++;
                mVisitorIsInsideComment = true;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called after all the child nodes of a Comment node have been visited.
            /// </summary>
            public override VisitorAction VisitCommentEnd(Comment comment)
            {
                mDocTraversalDepth--;
                IndentAndAppendLine("[Comment end]");
                mVisitorIsInsideComment = false;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Append a line to the StringBuilder, and indent it depending on how deep the visitor is
            /// into a comment/comment range's tree of child nodes.
            /// </summary>
            /// <param name="text"></param>
            private void IndentAndAppendLine(string text)
            {
                for (int i = 0; i < mDocTraversalDepth; i++)
                {
                    mBuilder.Append("|  ");
                }

                mBuilder.AppendLine(text);
            }

            private bool mVisitorIsInsideComment;
            private int mDocTraversalDepth;
            private readonly StringBuilder mBuilder;
        }
        //ExEnd

        private static void TestCommentsToText(CommentStructurePrinter visitor)
        {
            string visitorText = visitor.GetText();

            Assert.True(visitorText.Contains("[Comment range start]"));
            Assert.True(visitorText.Contains("[Comment range end]"));
            Assert.True(visitorText.Contains("[Comment start]"));
            Assert.True(visitorText.Contains("[Comment end]"));
            Assert.True(visitorText.Contains("[Run]"));
        }

        //ExStart
        //ExFor:DocumentVisitor.VisitFieldStart
        //ExFor:DocumentVisitor.VisitFieldEnd
        //ExFor:DocumentVisitor.VisitFieldSeparator
        //ExSummary:Shows how to print the node structure of every field in a document.
        [Test] //ExSkip
        public void FieldToText()
        {
            Document doc = new Document(MyDir + "DocumentVisitor-compatible features.docx");
            FieldStructurePrinter visitor = new FieldStructurePrinter();

            // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
            // and then traverses all the node's children in a depth-first manner.
            // The visitor can read and modify each visited node.
            doc.Accept(visitor);

            Console.WriteLine(visitor.GetText());
            TestFieldToText(visitor); //ExSkip
        }

        /// <summary>
        /// Traverses a node's non-binary tree of child nodes.
        /// Creates a map in the form of a string of all encountered Field nodes and their children.
        /// </summary>
        public class FieldStructurePrinter : DocumentVisitor
        {
            public FieldStructurePrinter()
            {
                mBuilder = new StringBuilder();
                mVisitorIsInsideField = false;
            }

            public string GetText()
            {
                return mBuilder.ToString();
            }

            /// <summary>
            /// Called when a Run node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitRun(Run run)
            {
                if (mVisitorIsInsideField) IndentAndAppendLine("[Run] \"" + run.GetText() + "\"");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FieldStart node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldStart(FieldStart fieldStart)
            {
                IndentAndAppendLine("[Field start] FieldType: " + fieldStart.FieldType);
                mDocTraversalDepth++;
                mVisitorIsInsideField = true;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FieldEnd node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldEnd(FieldEnd fieldEnd)
            {
                mDocTraversalDepth--;
                IndentAndAppendLine("[Field end]");
                mVisitorIsInsideField = false;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FieldSeparator node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldSeparator(FieldSeparator fieldSeparator)
            {
                IndentAndAppendLine("[FieldSeparator]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Append a line to the StringBuilder, and indent it depending on how deep the visitor is
            /// into the field's tree of child nodes.
            /// </summary>
            /// <param name="text"></param>
            private void IndentAndAppendLine(string text)
            {
                for (int i = 0; i < mDocTraversalDepth; i++)
                {
                    mBuilder.Append("|  ");
                }

                mBuilder.AppendLine(text);
            }

            private bool mVisitorIsInsideField;
            private int mDocTraversalDepth;
            private readonly StringBuilder mBuilder;
        }
        //ExEnd

        private static void TestFieldToText(FieldStructurePrinter visitor)
        {
            string visitorText = visitor.GetText();

            Assert.True(visitorText.Contains("[Field start]"));
            Assert.True(visitorText.Contains("[Field end]"));
            Assert.True(visitorText.Contains("[FieldSeparator]"));
            Assert.True(visitorText.Contains("[Run]"));
        }

        //ExStart
        //ExFor:DocumentVisitor.VisitHeaderFooterStart(HeaderFooter)
        //ExFor:DocumentVisitor.VisitHeaderFooterEnd(HeaderFooter)
        //ExFor:HeaderFooter.Accept(DocumentVisitor)
        //ExFor:HeaderFooterCollection.ToArray
        //ExFor:Run.Accept(DocumentVisitor)
        //ExFor:Run.GetText
        //ExSummary:Shows how to print the node structure of every header and footer in a document.
        [Test] //ExSkip
        public void HeaderFooterToText()
        {
            Document doc = new Document(MyDir + "DocumentVisitor-compatible features.docx");
            HeaderFooterStructurePrinter visitor = new HeaderFooterStructurePrinter();

            // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
            // and then traverses all the node's children in a depth-first manner.
            // The visitor can read and modify each visited node.
            doc.Accept(visitor);

            Console.WriteLine(visitor.GetText());

            // An alternative way of accessing a document's header/footers section-by-section is by accessing the collection.
            HeaderFooter[] headerFooters = doc.FirstSection.HeadersFooters.ToArray();
            Assert.AreEqual(3, headerFooters.Length);
            TestHeaderFooterToText(visitor); //ExSkip
        }

        /// <summary>
        /// Traverses a node's non-binary tree of child nodes.
        /// Creates a map in the form of a string of all encountered HeaderFooter nodes and their children.
        /// </summary>
        public class HeaderFooterStructurePrinter : DocumentVisitor
        {
            public HeaderFooterStructurePrinter()
            {
                mBuilder = new StringBuilder();
                mVisitorIsInsideHeaderFooter = false;
            }

            public string GetText()
            {
                return mBuilder.ToString();
            }

            /// <summary>
            /// Called when a Run node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitRun(Run run)
            {
                if (mVisitorIsInsideHeaderFooter) IndentAndAppendLine("[Run] \"" + run.GetText() + "\"");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a HeaderFooter node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitHeaderFooterStart(HeaderFooter headerFooter)
            {
                IndentAndAppendLine("[HeaderFooter start] HeaderFooterType: " + headerFooter.HeaderFooterType);
                mDocTraversalDepth++;
                mVisitorIsInsideHeaderFooter = true;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called after all the child nodes of a HeaderFooter node have been visited.
            /// </summary>
            public override VisitorAction VisitHeaderFooterEnd(HeaderFooter headerFooter)
            {
                mDocTraversalDepth--;
                IndentAndAppendLine("[HeaderFooter end]");
                mVisitorIsInsideHeaderFooter = false;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Append a line to the StringBuilder, and indent it depending on how deep the visitor is into the document tree.
            /// </summary>
            /// <param name="text"></param>
            private void IndentAndAppendLine(string text)
            {
                for (int i = 0; i < mDocTraversalDepth; i++) mBuilder.Append("|  ");

                mBuilder.AppendLine(text);
            }

            private bool mVisitorIsInsideHeaderFooter;
            private int mDocTraversalDepth;
            private readonly StringBuilder mBuilder;
        }
        //ExEnd

        private static void TestHeaderFooterToText(HeaderFooterStructurePrinter visitor)
        {
            string visitorText = visitor.GetText();

            Assert.True(visitorText.Contains("[HeaderFooter start] HeaderFooterType: HeaderPrimary"));
            Assert.True(visitorText.Contains("[HeaderFooter end]"));
            Assert.True(visitorText.Contains("[HeaderFooter start] HeaderFooterType: HeaderFirst"));
            Assert.True(visitorText.Contains("[HeaderFooter start] HeaderFooterType: HeaderEven"));
            Assert.True(visitorText.Contains("[HeaderFooter start] HeaderFooterType: FooterPrimary"));
            Assert.True(visitorText.Contains("[HeaderFooter start] HeaderFooterType: FooterFirst"));
            Assert.True(visitorText.Contains("[HeaderFooter start] HeaderFooterType: FooterEven"));
            Assert.True(visitorText.Contains("[Run]"));
        }

        //ExStart
        //ExFor:DocumentVisitor.VisitEditableRangeEnd(EditableRangeEnd)
        //ExFor:DocumentVisitor.VisitEditableRangeStart(EditableRangeStart)
        //ExSummary:Shows how to print the node structure of every editable range in a document.
        [Test] //ExSkip
        public void EditableRangeToText()
        {
            Document doc = new Document(MyDir + "DocumentVisitor-compatible features.docx");
            EditableRangeStructurePrinter visitor = new EditableRangeStructurePrinter();

            // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
            // and then traverses all the node's children in a depth-first manner.
            // The visitor can read and modify each visited node.
            doc.Accept(visitor);

            Console.WriteLine(visitor.GetText());
            TestEditableRangeToText(visitor); //ExSkip
        }

        /// <summary>
        /// Traverses a node's non-binary tree of child nodes.
        /// Creates a map in the form of a string of all encountered EditableRange nodes and their children.
        /// </summary>
        public class EditableRangeStructurePrinter : DocumentVisitor
        {
            public EditableRangeStructurePrinter()
            {
                mBuilder = new StringBuilder();
                mVisitorIsInsideEditableRange = false;
            }

            /// <summary>
            /// Gets the plain text of the document that was accumulated by the visitor.
            /// </summary>
            public string GetText()
            {
                return mBuilder.ToString();
            }

            /// <summary>
            /// Called when a Run node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitRun(Run run)
            {
                // We want to print the contents of runs, but only if they are inside shapes, as they would be in the case of text boxes
                if (mVisitorIsInsideEditableRange) IndentAndAppendLine("[Run] \"" + run.GetText() + "\"");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when an EditableRange node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitEditableRangeStart(EditableRangeStart editableRangeStart)
            {
                IndentAndAppendLine("[EditableRange start] ID: " + editableRangeStart.Id + " Owner: " +
                                    editableRangeStart.EditableRange.SingleUser);
                mDocTraversalDepth++;
                mVisitorIsInsideEditableRange = true;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a EditableRange node is ended.
            /// </summary>
            public override VisitorAction VisitEditableRangeEnd(EditableRangeEnd editableRangeEnd)
            {
                mDocTraversalDepth--;
                IndentAndAppendLine("[EditableRange end]");
                mVisitorIsInsideEditableRange = false;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
            /// </summary>
            /// <param name="text"></param>
            private void IndentAndAppendLine(string text)
            {
                for (int i = 0; i < mDocTraversalDepth; i++) mBuilder.Append("|  ");

                mBuilder.AppendLine(text);
            }

            private bool mVisitorIsInsideEditableRange;
            private int mDocTraversalDepth;
            private readonly StringBuilder mBuilder;
        }
        //ExEnd
        
        private static void TestEditableRangeToText(EditableRangeStructurePrinter visitor)
        {
            string visitorText = visitor.GetText();

            Assert.True(visitorText.Contains("[EditableRange start]"));
            Assert.True(visitorText.Contains("[EditableRange end]"));
            Assert.True(visitorText.Contains("[Run]"));
        }

        //ExStart
        //ExFor:DocumentVisitor.VisitFootnoteEnd(Footnote)
        //ExFor:DocumentVisitor.VisitFootnoteStart(Footnote)
        //ExFor:Footnote.Accept(DocumentVisitor)
        //ExSummary:Shows how to print the node structure of every footnote in a document.
        [Test] //ExSkip
        public void FootnoteToText()
        {
            Document doc = new Document(MyDir + "DocumentVisitor-compatible features.docx");
            FootnoteStructurePrinter visitor = new FootnoteStructurePrinter();

            // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
            // and then traverses all the node's children in a depth-first manner.
            // The visitor can read and modify each visited node.
            doc.Accept(visitor);

            Console.WriteLine(visitor.GetText());
            TestFootnoteToText(visitor); //ExSkip
        }

        /// <summary>
        /// Traverses a node's non-binary tree of child nodes.
        /// Creates a map in the form of a string of all encountered Footnote nodes and their children.
        /// </summary>
        public class FootnoteStructurePrinter : DocumentVisitor
        {
            public FootnoteStructurePrinter()
            {
                mBuilder = new StringBuilder();
                mVisitorIsInsideFootnote = false;
            }

            /// <summary>
            /// Gets the plain text of the document that was accumulated by the visitor.
            /// </summary>
            public string GetText()
            {
                return mBuilder.ToString();
            }

            /// <summary>
            /// Called when a Footnote node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFootnoteStart(Footnote footnote)
            {
                IndentAndAppendLine("[Footnote start] Type: " + footnote.FootnoteType);
                mDocTraversalDepth++;
                mVisitorIsInsideFootnote = true;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called after all the child nodes of a Footnote node have been visited.
            /// </summary>
            public override VisitorAction VisitFootnoteEnd(Footnote footnote)
            {
                mDocTraversalDepth--;
                IndentAndAppendLine("[Footnote end]");
                mVisitorIsInsideFootnote = false;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Run node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitRun(Run run)
            {
                if (mVisitorIsInsideFootnote) IndentAndAppendLine("[Run] \"" + run.GetText() + "\"");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
            /// </summary>
            /// <param name="text"></param>
            private void IndentAndAppendLine(string text)
            {
                for (int i = 0; i < mDocTraversalDepth; i++) mBuilder.Append("|  ");

                mBuilder.AppendLine(text);
            }

            private bool mVisitorIsInsideFootnote;
            private int mDocTraversalDepth;
            private readonly StringBuilder mBuilder;
        }
        //ExEnd

        private static void TestFootnoteToText(FootnoteStructurePrinter visitor)
        {
            string visitorText = visitor.GetText();

            Assert.True(visitorText.Contains("[Footnote start] Type: Footnote"));
            Assert.True(visitorText.Contains("[Footnote end]"));
            Assert.True(visitorText.Contains("[Run]"));
        }
        
        //ExStart
        //ExFor:DocumentVisitor.VisitOfficeMathEnd(Math.OfficeMath)
        //ExFor:DocumentVisitor.VisitOfficeMathStart(Math.OfficeMath)
        //ExFor:Math.MathObjectType
        //ExFor:Math.OfficeMath.Accept(DocumentVisitor)
        //ExFor:Math.OfficeMath.MathObjectType
        //ExSummary:Shows how to print the node structure of every office math node in a document.
        [Test] //ExSkip
        public void OfficeMathToText()
        {
            Document doc = new Document(MyDir + "DocumentVisitor-compatible features.docx");
            OfficeMathStructurePrinter visitor = new OfficeMathStructurePrinter();

            // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
            // and then traverses all the node's children in a depth-first manner.
            // The visitor can read and modify each visited node.
            doc.Accept(visitor);

            Console.WriteLine(visitor.GetText());
            TestOfficeMathToText(visitor); //ExSkip
        }

        /// <summary>
        /// Traverses a node's non-binary tree of child nodes.
        /// Creates a map in the form of a string of all encountered OfficeMath nodes and their children.
        /// </summary>
        public class OfficeMathStructurePrinter : DocumentVisitor
        {
            public OfficeMathStructurePrinter()
            {
                mBuilder = new StringBuilder();
                mVisitorIsInsideOfficeMath = false;
            }

            /// <summary>
            /// Gets the plain text of the document that was accumulated by the visitor.
            /// </summary>
            public string GetText()
            {
                return mBuilder.ToString();
            }

            /// <summary>
            /// Called when a Run node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitRun(Run run)
            {
                if (mVisitorIsInsideOfficeMath) IndentAndAppendLine("[Run] \"" + run.GetText() + "\"");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when an OfficeMath node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitOfficeMathStart(OfficeMath officeMath)
            {
                IndentAndAppendLine("[OfficeMath start] Math object type: " + officeMath.MathObjectType);
                mDocTraversalDepth++;
                mVisitorIsInsideOfficeMath = true;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called after all the child nodes of an OfficeMath node have been visited.
            /// </summary>
            public override VisitorAction VisitOfficeMathEnd(OfficeMath officeMath)
            {
                mDocTraversalDepth--;
                IndentAndAppendLine("[OfficeMath end]");
                mVisitorIsInsideOfficeMath = false;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
            /// </summary>
            /// <param name="text"></param>
            private void IndentAndAppendLine(string text)
            {
                for (int i = 0; i < mDocTraversalDepth; i++) mBuilder.Append("|  ");

                mBuilder.AppendLine(text);
            }

            private bool mVisitorIsInsideOfficeMath;
            private int mDocTraversalDepth;
            private readonly StringBuilder mBuilder;
        }
        //ExEnd

        private static void TestOfficeMathToText(OfficeMathStructurePrinter visitor)
        {
            string visitorText = visitor.GetText();

            Assert.True(visitorText.Contains("[OfficeMath start] Math object type: OMathPara"));
            Assert.True(visitorText.Contains("[OfficeMath start] Math object type: OMath"));
            Assert.True(visitorText.Contains("[OfficeMath start] Math object type: Argument"));
            Assert.True(visitorText.Contains("[OfficeMath start] Math object type: Supercript"));
            Assert.True(visitorText.Contains("[OfficeMath start] Math object type: SuperscriptPart"));
            Assert.True(visitorText.Contains("[OfficeMath start] Math object type: Fraction"));
            Assert.True(visitorText.Contains("[OfficeMath start] Math object type: Numerator"));
            Assert.True(visitorText.Contains("[OfficeMath start] Math object type: Denominator"));
            Assert.True(visitorText.Contains("[OfficeMath end]"));
            Assert.True(visitorText.Contains("[Run]"));
        }

        //ExStart
        //ExFor:DocumentVisitor.VisitSmartTagEnd(Markup.SmartTag)
        //ExFor:DocumentVisitor.VisitSmartTagStart(Markup.SmartTag)
        //ExSummary:Shows how to print the node structure of every smart tag in a document.
        [Test] //ExSkip
        public void SmartTagToText()
        {
            Document doc = new Document(MyDir + "Smart tags.doc");
            SmartTagStructurePrinter visitor = new SmartTagStructurePrinter();

            // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
            // and then traverses all the node's children in a depth-first manner.
            // The visitor can read and modify each visited node.
            doc.Accept(visitor);

            Console.WriteLine(visitor.GetText());
            TestSmartTagToText(visitor); //ExEnd
        }

        /// <summary>
        /// Traverses a node's non-binary tree of child nodes.
        /// Creates a map in the form of a string of all encountered SmartTag nodes and their children.
        /// </summary>
        public class SmartTagStructurePrinter : DocumentVisitor
        {
            public SmartTagStructurePrinter()
            {
                mBuilder = new StringBuilder();
                mVisitorIsInsideSmartTag = false;
            }

            /// <summary>
            /// Gets the plain text of the document that was accumulated by the visitor.
            /// </summary>
            public string GetText()
            {
                return mBuilder.ToString();
            }

            /// <summary>
            /// Called when a Run node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitRun(Run run)
            {
                if (mVisitorIsInsideSmartTag) IndentAndAppendLine("[Run] \"" + run.GetText() + "\"");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a SmartTag node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitSmartTagStart(SmartTag smartTag)
            {
                IndentAndAppendLine("[SmartTag start] Name: " + smartTag.Element);
                mDocTraversalDepth++;
                mVisitorIsInsideSmartTag = true;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called after all the child nodes of a SmartTag node have been visited.
            /// </summary>
            public override VisitorAction VisitSmartTagEnd(SmartTag smartTag)
            {
                mDocTraversalDepth--;
                IndentAndAppendLine("[SmartTag end]");
                mVisitorIsInsideSmartTag = false;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
            /// </summary>
            /// <param name="text"></param>
            private void IndentAndAppendLine(string text)
            {
                for (int i = 0; i < mDocTraversalDepth; i++) mBuilder.Append("|  ");

                mBuilder.AppendLine(text);
            }

            private bool mVisitorIsInsideSmartTag;
            private int mDocTraversalDepth;
            private readonly StringBuilder mBuilder;
        }
        //ExEnd

        private static void TestSmartTagToText(SmartTagStructurePrinter visitor)
        {
            string visitorText = visitor.GetText();

            Assert.True(visitorText.Contains("[SmartTag start] Name: address"));
            Assert.True(visitorText.Contains("[SmartTag start] Name: Street"));
            Assert.True(visitorText.Contains("[SmartTag start] Name: PersonName"));
            Assert.True(visitorText.Contains("[SmartTag start] Name: title"));
            Assert.True(visitorText.Contains("[SmartTag start] Name: GivenName"));
            Assert.True(visitorText.Contains("[SmartTag start] Name: Sn"));
            Assert.True(visitorText.Contains("[SmartTag start] Name: stockticker"));
            Assert.True(visitorText.Contains("[SmartTag start] Name: date"));
            Assert.True(visitorText.Contains("[SmartTag end]"));
            Assert.True(visitorText.Contains("[Run]"));
        }

        //ExStart
        //ExFor:StructuredDocumentTag.Accept(DocumentVisitor)
        //ExFor:DocumentVisitor.VisitStructuredDocumentTagEnd(Markup.StructuredDocumentTag)
        //ExFor:DocumentVisitor.VisitStructuredDocumentTagStart(Markup.StructuredDocumentTag)
        //ExSummary:Shows how to print the node structure of every structured document tag in a document.
        [Test] //ExSkip
        public void StructuredDocumentTagToText()
        {
            Document doc = new Document(MyDir + "DocumentVisitor-compatible features.docx");
            StructuredDocumentTagNodePrinter visitor = new StructuredDocumentTagNodePrinter();

            // When we get a composite node to accept a document visitor, the visitor visits the accepting node,
            // and then traverses all the node's children in a depth-first manner.
            // The visitor can read and modify each visited node.
            doc.Accept(visitor);

            Console.WriteLine(visitor.GetText());
            TestStructuredDocumentTagToText(visitor); //ExSkip
        }

        /// <summary>
        /// Traverses a node's non-binary tree of child nodes.
        /// Creates a map in the form of a string of all encountered StructuredDocumentTag nodes and their children.
        /// </summary>
        public class StructuredDocumentTagNodePrinter : DocumentVisitor
        {
            public StructuredDocumentTagNodePrinter()
            {
                mBuilder = new StringBuilder();
                mVisitorIsInsideStructuredDocumentTag = false;
            }

            /// <summary>
            /// Gets the plain text of the document that was accumulated by the visitor.
            /// </summary>
            public string GetText()
            {
                return mBuilder.ToString();
            }

            /// <summary>
            /// Called when a Run node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitRun(Run run)
            {
                if (mVisitorIsInsideStructuredDocumentTag) IndentAndAppendLine("[Run] \"" + run.GetText() + "\"");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a StructuredDocumentTag node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitStructuredDocumentTagStart(StructuredDocumentTag sdt)
            {
                IndentAndAppendLine("[StructuredDocumentTag start] Title: " + sdt.Title);
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called after all the child nodes of a StructuredDocumentTag node have been visited.
            /// </summary>
            public override VisitorAction VisitStructuredDocumentTagEnd(StructuredDocumentTag sdt)
            {
                mDocTraversalDepth--;
                IndentAndAppendLine("[StructuredDocumentTag end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
            /// </summary>
            /// <param name="text"></param>
            private void IndentAndAppendLine(string text)
            {
                for (int i = 0; i < mDocTraversalDepth; i++) mBuilder.Append("|  ");

                mBuilder.AppendLine(text);
            }

            private readonly bool mVisitorIsInsideStructuredDocumentTag;
            private int mDocTraversalDepth;
            private readonly StringBuilder mBuilder;
        }
        //ExEnd

        private static void TestStructuredDocumentTagToText(StructuredDocumentTagNodePrinter visitor)
        {
            string visitorText = visitor.GetText();

            Assert.True(visitorText.Contains("[StructuredDocumentTag start]"));
            Assert.True(visitorText.Contains("[StructuredDocumentTag end]"));
        }
    }
}