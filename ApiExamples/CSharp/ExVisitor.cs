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
using NUnit.Framework;

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
        //ExFor:DocumentVisitor.VisitParagraphEnd
        //ExFor:DocumentVisitor.VisitHeaderFooterStart
        //ExFor:DocumentVisitor.VisitBuildingBlockEnd(BuildingBlocks.BuildingBlock)
        //ExFor:DocumentVisitor.VisitBuildingBlockStart(BuildingBlocks.BuildingBlock)
        //ExFor:DocumentVisitor.VisitCellStart(Tables.Cell)
        //ExFor:DocumentVisitor.VisitCommentEnd(Comment)
        //ExFor:DocumentVisitor.VisitCommentRangeEnd(CommentRangeEnd)
        //ExFor:DocumentVisitor.VisitCommentRangeStart(CommentRangeStart)
        //ExFor:DocumentVisitor.VisitDocumentEnd(Document)
        //ExFor:DocumentVisitor.VisitDocumentStart(Document)
        //ExFor:DocumentVisitor.VisitEditableRangeEnd(EditableRangeEnd)
        //ExFor:DocumentVisitor.VisitEditableRangeStart(EditableRangeStart)
        //ExFor:DocumentVisitor.VisitFootnoteEnd(Footnote)
        //ExFor:DocumentVisitor.VisitGlossaryDocumentEnd(BuildingBlocks.GlossaryDocument)
        //ExFor:DocumentVisitor.VisitGlossaryDocumentStart(BuildingBlocks.GlossaryDocument)
        //ExFor:DocumentVisitor.VisitGroupShapeEnd(Drawing.GroupShape)
        //ExFor:DocumentVisitor.VisitHeaderFooterEnd(HeaderFooter)
        //ExFor:DocumentVisitor.VisitOfficeMathEnd(Math.OfficeMath)
        //ExFor:DocumentVisitor.VisitOfficeMathStart(Math.OfficeMath)
        //ExFor:DocumentVisitor.VisitRowStart(Tables.Row)
        //ExFor:DocumentVisitor.VisitSectionEnd(Section)
        //ExFor:DocumentVisitor.VisitSectionStart(Section)
        //ExFor:DocumentVisitor.VisitShapeEnd(Drawing.Shape)
        //ExFor:DocumentVisitor.VisitSmartTagEnd(Markup.SmartTag)
        //ExFor:DocumentVisitor.VisitSmartTagStart(Markup.SmartTag)
        //ExFor:DocumentVisitor.VisitStructuredDocumentTagEnd(Markup.StructuredDocumentTag)
        //ExFor:DocumentVisitor.VisitStructuredDocumentTagStart(Markup.StructuredDocumentTag)
        //ExFor:DocumentVisitor.VisitSubDocument(SubDocument)
        //ExFor:DocumentVisitor.VisitTableStart(Tables.Table)
        //ExFor:VisitorAction
        //ExId:ExtractContentDocToTxtConverter
        //ExSummary:Shows how to use the Visitor pattern to add new operations to the Aspose.Words object model. In this case we create a simple document converter into a text format.
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
        /// Simple implementation of saving a document in the plain text format. Implemented as a Visitor.
        /// </summary>
        public class MyDocToTxtWriter : DocumentVisitor
        {
            public MyDocToTxtWriter()
            {
                this.mIsSkipText = false;
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
                this.AppendText(run.Text);

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
                this.mIsSkipText = true;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FieldSeparator node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldSeparator(FieldSeparator fieldSeparator)
            {
                // Once reached a field separator node, we enable the output because we are
                // now entering the field result nodes.
                this.mIsSkipText = false;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FieldEnd node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldEnd(FieldEnd fieldEnd)
            {
                // Make sure we enable the output when reached a field end because some fields
                // do not have field separator and do not have field result.
                this.mIsSkipText = false;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when visiting of a Paragraph node is ended in the document.
            /// </summary>
            public override VisitorAction VisitParagraphEnd(Paragraph paragraph)
            {
                // When outputting to plain text we output Cr+Lf characters.
                this.AppendText(ControlChar.CrLf);

                return VisitorAction.Continue;
            }

            public override VisitorAction VisitBodyStart(Body body)
            {
                // We can detect beginning and end of all composite nodes such as Section, Body, 
                // Table, Paragraph etc and provide custom handling for them.
                this.mBuilder.Append("*** Body Started ***\r\n");

                return VisitorAction.Continue;
            }

            public override VisitorAction VisitBodyEnd(Body body)
            {
                this.mBuilder.Append("*** Body Ended ***\r\n");
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
            /// Called when an AbsolutePositionTab is encountered in the document.
            /// </summary>
            public override VisitorAction VisitAbsolutePositionTab(AbsolutePositionTab tab)
            {
                this.mBuilder.Append("\t");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a BookmarkStart is encountered in the document.
            /// </summary>
            public override VisitorAction VisitBookmarkStart(BookmarkStart bookmarkStart)
            {
                this.mIsSkipText = false;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a BookmarkEnd is encountered in the document.
            /// </summary>
            public override VisitorAction VisitBookmarkEnd(BookmarkEnd bookmarkEnd)
            {
                this.mIsSkipText = false;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a BuildingBlock is encountered in the document.
            /// </summary>
            public override VisitorAction VisitBuildingBlockStart(BuildingBlock buildingBlock)
            {
                this.mIsSkipText = false;
                this.mBuilder.Append(buildingBlock.GetText());

                return VisitorAction.Continue;
            }

            public override VisitorAction VisitBuildingBlockEnd(BuildingBlock buildingBlock)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitCellStart(Aspose.Words.Tables.Cell cell)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitCommentRangeStart(CommentRangeStart commentRangeStart)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitCommentRangeEnd(CommentRangeEnd commentRangeEnd)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitCommentEnd(Comment comment)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitDocumentStart(Document document)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitDocumentEnd(Document document)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitEditableRangeStart(EditableRangeStart editableRangeStart)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitEditableRangeEnd(EditableRangeEnd editableRangeEnd)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitFootnoteEnd(Footnote footnote)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitGlossaryDocumentStart(GlossaryDocument glossaryDocument)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitGlossaryDocumentEnd(GlossaryDocument glossaryDocument)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitGroupShapeEnd(GroupShape groupShape)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitHeaderFooterEnd(HeaderFooter headerFooter)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitOfficeMathStart(OfficeMath officeMath)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitOfficeMathEnd(OfficeMath officeMath)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitRowStart(Aspose.Words.Tables.Row row)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitSectionStart(Section section)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitSectionEnd(Section section)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitShapeEnd(Shape shape)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitSmartTagStart(SmartTag smartTag)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitSmartTagEnd(SmartTag smartTag)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitStructuredDocumentTagStart(StructuredDocumentTag smartTag)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitStructuredDocumentTagEnd(StructuredDocumentTag smartTag)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitSubDocument(SubDocument subDocument)
            {
                return VisitorAction.Continue;
            }

            public override VisitorAction VisitTableStart(Aspose.Words.Tables.Table table)
            {
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Adds text to the current output. Honours the enabled/disabled output flag.
            /// </summary>
            private void AppendText(String text)
            {
                if (!this.mIsSkipText)
                    this.mBuilder.Append(text);
            }

            private readonly StringBuilder mBuilder;
            private bool mIsSkipText;
        }
        //ExEnd
    }

    //TODO

}