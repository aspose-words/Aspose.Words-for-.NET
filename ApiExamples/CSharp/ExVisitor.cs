// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Text;

using Aspose.Words;
using Aspose.Words.Fields;

using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExVisitor : ApiExampleBase
    {
        /// <summary>
        /// This calls the below method to resolve skipping of [Test] in VB.NET.
        /// </summary>
        [Test]
        public void ToTextCaller()
        {
            this.ToText();
        }

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
        //ExFor:VisitorAction
        //ExId:ExtractContentDocToTxtConverter
        //ExSummary:Shows how to use the Visitor pattern to add new operations to the Aspose.Words object model. In this case we create a simple document converter into a text format.
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
            public string GetText()
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
                this.mBuilder.Append("[");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a BookmarkEnd is encountered in the document.
            /// </summary>
            public override VisitorAction VisitBookmarkEnd(BookmarkEnd bookmarkEnd)
            {
                this.mBuilder.Append("]");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Adds text to the current output. Honours the enabled/disabled output flag.
            /// </summary>
            private void AppendText(string text)
            {
                if (!this.mIsSkipText)
                    this.mBuilder.Append(text);
            }

            private readonly StringBuilder mBuilder;
            private bool mIsSkipText;
        }
        //ExEnd
    }
}