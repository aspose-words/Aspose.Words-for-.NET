using System.Text;
using NUnit.Framework;
using Aspose.Words;
using Aspose.Words.Fields;

namespace ApiExamples
{
    internal class ExAbsolutePositionTab : ApiExampleBase
    {
        [Test]
        //ExStart
        //ExFor:ExAbsolutePositionTab
        //ExFor:ExAbsolutePositionTab.Accept(DocumentVisitor)
        //ExSummary:Shows how to use AbsolutePositionTab.
        public void DocumentToTxt()
        {         
            // This document contains two sentences separated by an absolute position tab.
            Document doc = new Document(MyDir + "AbsolutePositionTab.docx");

            // An AbsolutePositionTab is a child node of a paragraph. 
            // It gets picked up when looking for nodes of the SpecialChar type.
            Paragraph para = doc.FirstSection.Body.FirstParagraph;
            AbsolutePositionTab absPositionTab = (AbsolutePositionTab)para.GetChild(NodeType.SpecialChar, 0, true);

            // This implementation of the DocumentVisitor pattern converts the document to plain text.
            MyDocToTxtWriter myDocToTxtWriter = new MyDocToTxtWriter();

            // We can run the DocumentVisitor over the whole paragraph.
            para.Accept(myDocToTxtWriter);

            // Tab character is placed where the AbsolutePositionTab was.
            Assert.AreEqual("Before AbsolutePositionTab\tAfter AbsolutePositionTab\r\n", myDocToTxtWriter.GetText());

            // An AbsolutePositionTab can accept a DocumentVisitor by itself too.
            myDocToTxtWriter = new MyDocToTxtWriter();
            absPositionTab.Accept(myDocToTxtWriter);

            Assert.AreEqual("\t", myDocToTxtWriter.GetText());
        }

        /// <summary>
        /// Simple implementation of saving a document in the plain text format. Implemented as a Visitor.
        /// </summary>
        public class MyDocToTxtWriter : DocumentVisitor
        {
            public MyDocToTxtWriter()
            {
                mIsSkipText = false;
                mBuilder = new StringBuilder();
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
                AppendText(run.Text);

                // Let the visitor continue visiting other nodes.
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when an AbsolutePositionTab node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitAbsolutePositionTab(AbsolutePositionTab tab)
            {
                // We'll treat the AbsolutePositionTab as a simple tab in this case.
                mBuilder.Append("\t");

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
                mIsSkipText = true;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FieldSeparator node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldSeparator(FieldSeparator fieldSeparator)
            {
                // Once reached a field separator node, we enable the output because we are
                // now entering the field result nodes.
                mIsSkipText = false;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when a FieldEnd node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitFieldEnd(FieldEnd fieldEnd)
            {
                // Make sure we enable the output when reached a field end because some fields
                // do not have field separator and do not have field result.
                mIsSkipText = false;

                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when visiting of a Paragraph node is ended in the document.
            /// </summary>
            public override VisitorAction VisitParagraphEnd(Paragraph paragraph)
            {
                // When outputting to plain text we output Cr+Lf characters.
                AppendText(ControlChar.CrLf);

                return VisitorAction.Continue;
            }

            public override VisitorAction VisitBodyStart(Body body)
            {
                // We can detect beginning and end of all composite nodes such as Section, Body, 
                // Table, Paragraph etc and provide custom handling for them.
                mBuilder.Append("*** Body Started ***\r\n");

                return VisitorAction.Continue;
            }

            public override VisitorAction VisitBodyEnd(Body body)
            {
                mBuilder.Append("*** Body Ended ***\r\n");
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
            /// Adds text to the current output. Honours the enabled/disabled output flag.
            /// </summary>
            private void AppendText(string text)
            {
                if (!mIsSkipText)
                    mBuilder.Append(text);
            }

            private readonly StringBuilder mBuilder;
            private bool mIsSkipText;
        }
        //ExEnd
    }
}
