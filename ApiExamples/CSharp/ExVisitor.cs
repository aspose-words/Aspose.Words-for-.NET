// Copyright (c) 2001-2017 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Text;
using Aspose.Words;
using Aspose.Words.BuildingBlocks;
using Aspose.Words.Drawing;
using Aspose.Words.Fields;
using Aspose.Words.Markup;
using Aspose.Words.Math;
using Aspose.Words.Tables;
using NUnit.Framework;
using Document = Aspose.Words.Document;
using HeaderFooter = Aspose.Words.HeaderFooter;

namespace ApiExamples
{
    [TestFixture]
    public class ExVisitor : ApiExampleBase
    {
        //ExStart
        //ExFor:Document.Accept
        //ExFor:Body.Accept
        //ExFor:DocumentVisitor
        //ExFor:DocumentVisitor.VisitAbsolutePositionTab
        //ExFor:DocumentVisitor.VisitBookmarkStart 
        //ExFor:DocumentVisitor.VisitBookmarkEnd
        //ExFor:DocumentVisitor.VisitRun
        //ExFor:DocumentVisitor.VisitFieldStart
        //ExFor:DocumentVisitor.VisitFieldEnd
        //ExFor:DocumentVisitor.VisitFieldSeparator
        //ExFor:DocumentVisitor.VisitBodyStart
        //ExFor:DocumentVisitor.VisitBodyEnd
        //ExFor:DocumentVisitor.VisitParagraphStart
        //ExFor:DocumentVisitor.VisitParagraphEnd
        //ExFor:DocumentVisitor.VisitHeaderFooterStart
        //ExFor:DocumentVisitor.VisitHeaderFooterEnd
        //ExFor:DocumentVisitor.VisitBuildingBlockEnd(BuildingBlocks.BuildingBlock)
        //ExFor:DocumentVisitor.VisitBuildingBlockStart(BuildingBlocks.BuildingBlock)
        //ExFor:DocumentVisitor.VisitCellStart(Tables.Cell)
        //ExFor:DocumentVisitor.VisitCellEnd(Tables.Cell)
        //ExFor:DocumentVisitor.VisitCommentStart(Comment)
        //ExFor:DocumentVisitor.VisitCommentEnd(Comment)
        //ExFor:DocumentVisitor.VisitCommentRangeEnd(CommentRangeEnd)
        //ExFor:DocumentVisitor.VisitCommentRangeStart(CommentRangeStart)
        //ExFor:DocumentVisitor.VisitDocumentEnd(Document)
        //ExFor:DocumentVisitor.VisitDocumentStart(Document)
        //ExFor:DocumentVisitor.VisitEditableRangeEnd(EditableRangeEnd)
        //ExFor:DocumentVisitor.VisitEditableRangeStart(EditableRangeStart)
        //ExFor:DocumentVisitor.VisitFootnoteEnd(Footnote)
        //ExFor:DocumentVisitor.VisitFootnoteStart(Footnote)
        //ExFor:DocumentVisitor.VisitGlossaryDocumentEnd(BuildingBlocks.GlossaryDocument)
        //ExFor:DocumentVisitor.VisitGlossaryDocumentStart(BuildingBlocks.GlossaryDocument)
        //ExFor:DocumentVisitor.VisitGroupShapeEnd(Drawing.GroupShape)
        //ExFor:DocumentVisitor.VisitGroupShapeStart(Drawing.GroupShape)
        //ExFor:DocumentVisitor.VisitHeaderFooterEnd(HeaderFooter)
        //ExFor:DocumentVisitor.VisitOfficeMathEnd(Math.OfficeMath)
        //ExFor:DocumentVisitor.VisitOfficeMathStart(Math.OfficeMath)
        //ExFor:DocumentVisitor.VisitRowEnd(Tables.Row)
        //ExFor:DocumentVisitor.VisitRowStart(Tables.Row)
        //ExFor:DocumentVisitor.VisitSectionEnd(Section)
        //ExFor:DocumentVisitor.VisitSectionStart(Section)
        //ExFor:DocumentVisitor.VisitShapeEnd(Drawing.Shape)
        //ExFor:DocumentVisitor.VisitShapeStart(Drawing.Shape)
        //ExFor:DocumentVisitor.VisitSmartTagEnd(Markup.SmartTag)
        //ExFor:DocumentVisitor.VisitSmartTagStart(Markup.SmartTag)
        //ExFor:DocumentVisitor.VisitStructuredDocumentTagEnd(Markup.StructuredDocumentTag)
        //ExFor:DocumentVisitor.VisitStructuredDocumentTagStart(Markup.StructuredDocumentTag)
        //ExFor:DocumentVisitor.VisitSubDocument(SubDocument)
        //ExFor:DocumentVisitor.VisitTableEnd(Tables.Table)
        //ExFor:DocumentVisitor.VisitTableStart(Tables.Table)
        //ExFor:VisitorAction
        //ExId:ExtractContentDocToTxtConverter
        //ExSummary:Shows how to use a visitor to traverse and interact with nodes in a document. In this case we are using a visitor that prints a directory tree-style map of a document.
        [Test] //ExSkip
        public void ToText()
        {
            // Open the document we want to map.
            Document doc = new Document(MyDir + "Visitor.MapDocument.doc");

            // Create an object that inherits from the DocumentVisitor class.
            DocToNodeTree converter = new DocToNodeTree();

            // This is the well known Visitor pattern. Get the model to accept a visitor.
            // The model will iterate through itself by calling the corresponding methods
            // on the visitor object (this is called visiting).
            // 
            // Note that every node in the object model has the Accept method so the visiting
            // can be executed not only for the whole document, but for any node in the document.
            doc.Accept(converter);

            // Once the visiting is complete, we can retrieve the result of the operation,
            // that in this example, has accumulated in the visitor.
            Console.WriteLine(converter.GetText());
        }

        /// <summary>
        /// This Visitor implementation traverses a document and maps it's nodes in the style of a vertical directory tree diagram. 
        /// </summary>
        public class DocToNodeTree : DocumentVisitor
        {
            public DocToNodeTree()
            {
                this.mBuilder = new StringBuilder();
            }

            /// <summary>
            /// Gets the plain text of the document that was accumulated by the visitor.
            /// </summary>
            public String GetText()
            {
                return this.mBuilder.ToString();
            }

            /// <summary>
            /// Called when a Document node is encountered.
            /// </summary>
            public override VisitorAction VisitDocumentStart(Document document)
            {
                int childNodeCount = document.GetChildNodes(NodeType.Any, true).Count;

                // A Document node is at the root of every document, so if we let a document accept a visitor, this will be the first visitor action to be carried out
                this.IndentAndAppendLine("[Document start] Child nodes: " + childNodeCount);
                mDocTraversalDepth++;

                // Let the visitor continue visiting other nodes
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a Document is ended.
            /// </summary>
            public override VisitorAction VisitDocumentEnd(Document document)
            {
                // If we let a document accept a visitor, this will be the last visitor action to be carried out
                mDocTraversalDepth--;
                this.IndentAndAppendLine("[Document end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Section node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitSectionStart(Section section)
            {
                // Get the index of our section within the document
                NodeCollection docSections = section.Document.GetChildNodes(NodeType.Section, false);
                int sectionIndex = docSections.IndexOf(section);

                this.IndentAndAppendLine("[Section start] Section index: " + sectionIndex);
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a Section node is ended.
            /// </summary>
            public override VisitorAction VisitSectionEnd(Section section)
            {
                mDocTraversalDepth--;
                this.IndentAndAppendLine("[Section end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Body node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitBodyStart(Body body)
            {
                int paragraphCount = body.Paragraphs.Count;
                this.IndentAndAppendLine("[Body start] Paragraphs: " + paragraphCount);
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a Body node is ended.
            /// </summary>
            public override VisitorAction VisitBodyEnd(Body body)
            {
                mDocTraversalDepth--;
                this.IndentAndAppendLine("[Body end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Paragraph node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitParagraphStart(Paragraph paragraph)
            {
                this.IndentAndAppendLine("[Paragraph start]");
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a Paragraph node is ended.
            /// </summary>
            public override VisitorAction VisitParagraphEnd(Paragraph paragraph)
            {
                mDocTraversalDepth--;
                this.IndentAndAppendLine("[Paragraph end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Run node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitRun(Run run)
            {
                // A Run is not a composite node; it has no children
                // Our visitor will not be going any deeper, so we'll just print the contents of the run on the same line
                this.IndentAndAppendLine("[Run] \"" + run.Text + "\"");

                return VisitorAction.Continue;
            }

            // We can expect every document to have Document, Section, Body, Paragraph and Run nodes, and we are handling all of those above
            // All other visitor actions below are optional, depending on the variety of nodes in the document

            /// <summary>
            /// Called when a FieldStart node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldStart(FieldStart fieldStart)
            {
                this.IndentAndAppendLine("[Field start] FieldType: " + fieldStart.FieldType);
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FieldEnd node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldEnd(FieldEnd fieldEnd)
            {
                mDocTraversalDepth--;
                this.IndentAndAppendLine("[Field end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FieldSeparator node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldSeparator(FieldSeparator fieldSeparator)
            {
                this.IndentAndAppendLine("[FieldSeparator]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a HeaderFooter node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitHeaderFooterStart(HeaderFooter headerFooter)
            {
                this.IndentAndAppendLine("[HeaderFooter start] HeaderFooterType: " + headerFooter.HeaderFooterType);
                mDocTraversalDepth++;

                // Next, the visitor will traverse the nodes in the header/footer
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a HeaderFooter node is ended.
            /// </summary>
            public override VisitorAction VisitHeaderFooterEnd(HeaderFooter headerFooter)
            {
                mDocTraversalDepth--;
                this.IndentAndAppendLine("[HeaderFooter end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when an AbsolutePositionTab is encountered in the document.
            /// </summary>
            public override VisitorAction VisitAbsolutePositionTab(AbsolutePositionTab tab)
            {
                // If we encounter an AbsolutePositionTab character, in our text output we can simply sibstitute it with a tab 
                this.mBuilder.Append("\t");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a BookmarkStart node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitBookmarkStart(BookmarkStart bookmarkStart)
            {
                this.IndentAndAppendLine("[Bookmark start] Name: \"" + bookmarkStart.Bookmark.Name + "\"");
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a BookmarkEnd node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitBookmarkEnd(BookmarkEnd bookmarkEnd)
            {
                mDocTraversalDepth--;
                this.IndentAndAppendLine("[Bookmark end]");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a BuildingBlock node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitBuildingBlockStart(BuildingBlock buildingBlock)
            {
                this.IndentAndAppendLine("[Building block] Name: " + buildingBlock.Name + "GUID: " + buildingBlock.Guid);
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a BuildingBlock node is ended in the document.
            /// </summary>
            public override VisitorAction VisitBuildingBlockEnd(BuildingBlock buildingBlock)
            {
                mDocTraversalDepth--;
                this.IndentAndAppendLine("[BuildingBlock end]");

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

                this.IndentAndAppendLine("[Table start] Size: " + rows + "x" + columns);
                mDocTraversalDepth++;

                // The visitor will now go through every row and cell
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a Table node is ended.
            /// </summary>
            public override VisitorAction VisitTableEnd(Table table)
            {
                mDocTraversalDepth--;
                this.IndentAndAppendLine("[Table end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Row node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitRowStart(Row row)
            {
                this.IndentAndAppendLine("[Row start]");
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a Row node is ended.
            /// </summary>
            public override VisitorAction VisitRowEnd(Row row)
            {
                mDocTraversalDepth--;
                this.IndentAndAppendLine("[Row end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Cell node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitCellStart(Cell cell)
            {
                Row row = cell.ParentRow;
                Table table = row.ParentTable;

                this.IndentAndAppendLine("[Cell start] Row " + (table.IndexOf(row) + 1) + ", Col " + (row.IndexOf(cell) + 1) + "");
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a Cell node is ended in the document.
            /// </summary>
            public override VisitorAction VisitCellEnd(Cell cell)
            {
                mDocTraversalDepth--;
                this.IndentAndAppendLine("[Cell end]");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a CommentRangeStart node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitCommentRangeStart(CommentRangeStart commentRangeStart)
            {
                this.IndentAndAppendLine("[Comment range start]");
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a CommentRangeEnd node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitCommentRangeEnd(CommentRangeEnd commentRangeEnd)
            {
                mDocTraversalDepth--;
                this.IndentAndAppendLine("[Comment range end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Comment node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitCommentStart(Comment comment)
            {
                this.IndentAndAppendLine(String.Format("[Comment start] {0}, {1}", comment.Author, comment.DateTime));
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a Comment node is ended in the document.
            /// </summary>
            public override VisitorAction VisitCommentEnd(Comment comment)
            {
                mDocTraversalDepth--;
                this.IndentAndAppendLine("[Comment end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when an EditableRange node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitEditableRangeStart(EditableRangeStart editableRangeStart)
            {
                this.IndentAndAppendLine("[EditableRange start]");
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a EditableRange node is ended.
            /// </summary>
            public override VisitorAction VisitEditableRangeEnd(EditableRangeEnd editableRangeEnd)
            {
                mDocTraversalDepth--;
                this.IndentAndAppendLine("[EditableRange end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Footnote node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFootnoteStart(Footnote footnote)
            {   
                this.IndentAndAppendLine("[Footnote start] Type: " + footnote.FootnoteType);
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a Footnote node is ended.
            /// </summary>
            public override VisitorAction VisitFootnoteEnd(Footnote footnote)
            {
                mDocTraversalDepth--;
                this.IndentAndAppendLine("[Footnote end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a GlossaryDocument node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitGlossaryDocumentStart(GlossaryDocument glossaryDocument)
            {
                this.IndentAndAppendLine("[GlossaryDocument start] Building block count: " + glossaryDocument.BuildingBlocks.Count);
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a GlossaryDocument node is ended.
            /// </summary>
            public override VisitorAction VisitGlossaryDocumentEnd(GlossaryDocument glossaryDocument)
            {
                mDocTraversalDepth--;
                this.IndentAndAppendLine("[GlossaryDocument end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a GroupShape node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitGroupShapeStart(GroupShape groupShape)
            {
                int innerShapeCount = groupShape.GetChildNodes(NodeType.Shape, true).Count;

                this.IndentAndAppendLine("[GroupShape start] Inner shapes: " + innerShapeCount);
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a GroupShape node is ended.
            /// </summary>
            public override VisitorAction VisitGroupShapeEnd(GroupShape groupShape)
            {
                mDocTraversalDepth--;
                this.IndentAndAppendLine("[GroupShape end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when an OfficeMath node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitOfficeMathStart(OfficeMath officeMath)
            {
                this.IndentAndAppendLine("[OfficeMath start] Math object type: " + officeMath.MathObjectType);
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a OfficeMath node is ended.
            /// </summary>
            public override VisitorAction VisitOfficeMathEnd(OfficeMath officeMath)
            {
                mDocTraversalDepth--;
                this.IndentAndAppendLine("[OfficeMath end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Shape node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitShapeStart(Shape shape)
            {
                this.IndentAndAppendLine("[Shape start] Type: " + shape.ShapeType + ", Fill color: " + shape.FillColor);
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a Shape node is ended.
            /// </summary>
            public override VisitorAction VisitShapeEnd(Shape shape)
            {
                mDocTraversalDepth--;
                this.IndentAndAppendLine("[Shape end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a SmartTag node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitSmartTagStart(SmartTag smartTag)
            {
                this.IndentAndAppendLine("[SmartTag start] Name: " + smartTag.Element);
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a SmartTag node is ended.
            /// </summary>
            public override VisitorAction VisitSmartTagEnd(SmartTag smartTag)
            {
                mDocTraversalDepth--;
                this.IndentAndAppendLine("[SmartTag end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a StructuredDocumentTag node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitStructuredDocumentTagStart(StructuredDocumentTag structuredDocumentTag)
            {
                this.IndentAndAppendLine("[StructuredDocumentTag start] Title: " + structuredDocumentTag.Title);
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a StructuredDocumentTag node is ended.
            /// </summary>
            public override VisitorAction VisitStructuredDocumentTagEnd(StructuredDocumentTag structuredDocumentTag)
            {
                mDocTraversalDepth--;
                this.IndentAndAppendLine("[StructuredDocumentTag end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a SubDocument node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitSubDocument(SubDocument subDocument)
            {
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Append a line to the StringBuilder and indent it depending on how deep the visitor is into the document tree.
            /// </summary>
            /// <param name="text"></param>
            private void IndentAndAppendLine(String text)
            {
                for (int i = 0; i < mDocTraversalDepth; i++)
                {
                    mBuilder.Append("|  ");
                }
                mBuilder.AppendLine(text);
            }

            private readonly StringBuilder mBuilder;
            private int mDocTraversalDepth;
        }
        //ExEnd
    }
}
