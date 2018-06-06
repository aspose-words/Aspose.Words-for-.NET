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
        //ExSummary:Shows how to use the Visitor pattern to add new operations to the Aspose.Words object model. In this case we accept a Visitor that prints a map of our document's tree of nodes to the console.
        [Test] //ExSkip
        public void ToText()
        {
            // Open the document we want to convert.
            Document doc = new Document(MyDir + "Visitor.ToText.doc");

            // Create an object that inherits from the DocumentVisitor class.
            MyDocToTxtWriter myConverter = new MyDocToTxtWriter();

            // This is the well known Visitor pattern. Get the model to accept a visitor.
            // The model will iterate through itself by calling the corresponding methods
            // on the visitor object (this is called visiting).
            // 
            // Note that every node in the object model has the Accept method so the visiting
            // can be executed not only for the whole document, but for any node in the document.
            doc.Accept(myConverter);

            // Once the visiting is complete, we can retrieve the result of the operation,
            // that in this example, has accumulated in the visitor.
            Console.WriteLine(myConverter.GetText());
        }

        /// <summary>
        /// Simple implementation of printing a document as a hierarchy of nodes. Implemented as a Visitor.
        /// </summary>
        public class MyDocToTxtWriter : DocumentVisitor
        {
            public MyDocToTxtWriter()
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
            /// Called when a Run node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitRun(Run run)
            {
                this.AppendIndentedLine("[Run \"" + run.Text + "\"]");

                // Let the visitor continue visiting other nodes.
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FieldStart node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldStart(FieldStart fieldStart)
            {
                // In Microsoft Word, a field code (such as "MERGEFIELD FieldName") follows
                // after a field start character. We want to skip field codes and output field 
                // result only, therefore we use a flag to suspend the output while inside a field code.
                //
                // Note this is a very simplistic implementation and will not work very well
                // if you have nested fields in a document. 
                this.AppendIndentedLine("[Field start]");
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FieldEnd node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldEnd(FieldEnd fieldEnd)
            {
                // Make sure we enable the output when reached a field end because some fields
                // do not have field separator and do not have field result.
                mDocTraversalDepth--;
                this.AppendIndentedLine("[Field end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FieldSeparator node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldSeparator(FieldSeparator fieldSeparator)
            {
                // Once reached a field separator node, we enable the output because we are
                // now entering the field result nodes.
                this.AppendIndentedLine("[FieldSeparator]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Paragraph node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitParagraphStart(Paragraph paragraph)
            {
                // When outputting to plain text we output Cr+Lf characters.
                this.AppendIndentedLine("[Paragraph start]");
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a Paragraph node is ended.
            /// </summary>
            public override VisitorAction VisitParagraphEnd(Paragraph paragraph)
            {
                // When outputting to plain text we output Cr+Lf characters.
                mDocTraversalDepth--;
                this.AppendIndentedLine("[Paragraph end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Body node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitBodyStart(Body body)
            {
                // We can detect beginning and end of all composite nodes such as Section, Body, 
                // Table, Paragraph etc and provide custom handling for them.
                this.AppendIndentedLine("[Body start]");
                mDocTraversalDepth++;
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a Body node is ended.
            /// </summary>
            public override VisitorAction VisitBodyEnd(Body body)
            {
                mDocTraversalDepth--;
                this.AppendIndentedLine("[Body end]");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a HeaderFooter node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitHeaderFooterStart(HeaderFooter headerFooter)
            {
                // Returning this value from a visitor method causes visiting of this
                // node to stop and move on to visiting the next sibling node.
                // The net effect in this example is that the text of headers and footers
                // is not included in the resulting output.
                return VisitorAction.SkipThisNode;
            }

            /// <summary>
            /// Called when the visiting of a HeaderFooter node is ended.
            /// </summary>
            public override VisitorAction VisitHeaderFooterEnd(HeaderFooter headerFooter)
            {
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when an AbsolutePositionTab is encountered in the document.
            /// </summary>
            public override VisitorAction VisitAbsolutePositionTab(AbsolutePositionTab tab)
            {
                this.mBuilder.Append("\t");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a BookmarkStart node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitBookmarkStart(BookmarkStart bookmarkStart)
            {
                Bookmark bookmark = bookmarkStart.Bookmark;
                this.AppendIndentedLine("[Bookmark Name: " + bookmark.Name + "]");
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a BookmarkEnd node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitBookmarkEnd(BookmarkEnd bookmarkEnd)
            {
                mDocTraversalDepth--;
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a BuildingBlock node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitBuildingBlockStart(BuildingBlock buildingBlock)
            {
                this.AppendIndentedLine("[Building block]");
                mDocTraversalDepth++;
                this.AppendIndentedLine("Name: " + buildingBlock.Name);
                this.AppendIndentedLine("GUID: " + buildingBlock.Guid);

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a BuildingBlock node is ended in the document.
            /// </summary>
            public override VisitorAction VisitBuildingBlockEnd(BuildingBlock buildingBlock)
            {
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Table is encountered in the document.
            /// </summary>
            public override VisitorAction VisitTableStart(Table table)
            {
                this.AppendIndentedLine("[Table start]");
                mDocTraversalDepth++;
                // At this point we could traverse the table and print the content of every cell...
                // But it would be more elegant to let the visitor carry on

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a Table node is ended.
            /// </summary>
            public override VisitorAction VisitTableEnd(Table table)
            {
                mDocTraversalDepth--;
                this.AppendIndentedLine("[Table end]");

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Row node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitRowStart(Row row)
            {
                this.AppendIndentedLine("[Row start]");
                mDocTraversalDepth++;
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a Row node is ended.
            /// </summary>
            public override VisitorAction VisitRowEnd(Row row)
            {
                mDocTraversalDepth--;
                this.AppendIndentedLine("[Row end]");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Cell node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitCellStart(Cell cell)
            {
                Row row = cell.ParentRow;
                Table table = row.ParentTable;

                this.AppendIndentedLine("[Cell in row " + (table.IndexOf(row) + 1) + ", column " + (row.IndexOf(cell) + 1) + "]");
                mDocTraversalDepth++;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a Cell node is ended in the document.
            /// </summary>
            public override VisitorAction VisitCellEnd(Cell cell)
            {
                mDocTraversalDepth--;
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a CommentRangeStart node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitCommentRangeStart(CommentRangeStart commentRangeStart)
            {
                this.AppendIndentedLine("[Comment range start]");
                mDocTraversalDepth++;
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a CommentRangeEnd node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitCommentRangeEnd(CommentRangeEnd commentRangeEnd)
            {
                mDocTraversalDepth--;
                this.AppendIndentedLine("[Comment range end]");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Comment node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitCommentStart(Comment comment)
            {
                this.AppendIndentedLine(String.Format("[Comment start] {0}, {1}", comment.Author, comment.DateTime));
                mDocTraversalDepth++;
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a Comment node is ended in the document.
            /// </summary>
            public override VisitorAction VisitCommentEnd(Comment comment)
            {
                mDocTraversalDepth--;
                this.AppendIndentedLine("[Comment end]");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Document node is encountered.
            /// </summary>
            public override VisitorAction VisitDocumentStart(Document document)
            {
                this.AppendIndentedLine("[Document start]");
                mDocTraversalDepth++;
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a Document is ended.
            /// </summary>
            public override VisitorAction VisitDocumentEnd(Document document)
            {
                mDocTraversalDepth--;
                this.AppendIndentedLine("[Document end]");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when an EditableRange node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitEditableRangeStart(EditableRangeStart editableRangeStart)
            {
                this.AppendIndentedLine("[EditableRange start]");
                mDocTraversalDepth++;
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a EditableRange node is ended.
            /// </summary>
            public override VisitorAction VisitEditableRangeEnd(EditableRangeEnd editableRangeEnd)
            {
                mDocTraversalDepth--;
                this.AppendIndentedLine("[EditableRange end]");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Footnote node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFootnoteStart(Footnote footnote)
            {
                this.AppendIndentedLine("[Footnote start]");
                mDocTraversalDepth++;
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a Footnote node is ended.
            /// </summary>
            public override VisitorAction VisitFootnoteEnd(Footnote footnote)
            {
                mDocTraversalDepth--;
                this.AppendIndentedLine("[Footnote end]");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a GlossaryDocument node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitGlossaryDocumentStart(GlossaryDocument glossaryDocument)
            {
                this.AppendIndentedLine("[GlossaryDocument start]");
                mDocTraversalDepth++;
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a GlossaryDocument node is ended.
            /// </summary>
            public override VisitorAction VisitGlossaryDocumentEnd(GlossaryDocument glossaryDocument)
            {
                mDocTraversalDepth--;
                this.AppendIndentedLine("[GlossaryDocument end]");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a GroupShape node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitGroupShapeStart(GroupShape groupShape)
            {
                this.AppendIndentedLine("[GroupShape start]");
                mDocTraversalDepth++;
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a GroupShape node is ended.
            /// </summary>
            public override VisitorAction VisitGroupShapeEnd(GroupShape groupShape)
            {
                mDocTraversalDepth--;
                this.AppendIndentedLine("[GroupShape end]");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when an OfficeMath node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitOfficeMathStart(OfficeMath officeMath)
            {
                this.AppendIndentedLine("[OfficeMath start]");
                mDocTraversalDepth++;
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a OfficeMath node is ended.
            /// </summary>
            public override VisitorAction VisitOfficeMathEnd(OfficeMath officeMath)
            {
                mDocTraversalDepth--;
                this.AppendIndentedLine("[OfficeMath end]");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Section node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitSectionStart(Section section)
            {
                this.AppendIndentedLine("[Section start]");
                mDocTraversalDepth++;
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a Section node is ended.
            /// </summary>
            public override VisitorAction VisitSectionEnd(Section section)
            {
                mDocTraversalDepth--;
                this.AppendIndentedLine("[Section end]");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a Shape node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitShapeStart(Shape shape)
            {
                this.AppendIndentedLine("[Shape start]");
                mDocTraversalDepth++;
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a Shape node is ended.
            /// </summary>
            public override VisitorAction VisitShapeEnd(Shape shape)
            {
                mDocTraversalDepth--;
                this.AppendIndentedLine("[Shape end]");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a SmartTag node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitSmartTagStart(SmartTag smartTag)
            {
                this.AppendIndentedLine("[SmartTag start]");
                mDocTraversalDepth++;
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a SmartTag node is ended.
            /// </summary>
            public override VisitorAction VisitSmartTagEnd(SmartTag smartTag)
            {
                mDocTraversalDepth--;
                this.AppendIndentedLine("[SmartTag end]");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a StructuredDocumentTag node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitStructuredDocumentTagStart(StructuredDocumentTag smartTag)
            {
                this.AppendIndentedLine("[StructuredDocumentTag start]");
                mDocTraversalDepth++;
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when the visiting of a StructuredDocumentTag node is ended.
            /// </summary>
            public override VisitorAction VisitStructuredDocumentTagEnd(StructuredDocumentTag smartTag)
            {
                mDocTraversalDepth--;
                this.AppendIndentedLine("[StructuredDocumentTag end]");
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
            private void AppendIndentedLine(String text)
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
