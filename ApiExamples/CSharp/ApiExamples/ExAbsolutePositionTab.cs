// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Text;
using NUnit.Framework;
using Aspose.Words;

namespace ApiExamples
{
    [TestFixture]
    public class ExAbsolutePositionTab : ApiExampleBase
    {
        //ExStart
        //ExFor:AbsolutePositionTab
        //ExFor:AbsolutePositionTab.Accept(DocumentVisitor)
        //ExFor:DocumentVisitor.VisitAbsolutePositionTab
        //ExSummary:Shows how to work with AbsolutePositionTab.
        [Test] //ExSkip
        public void DocumentToTxt()
        {
            // This document contains two sentences separated by an absolute position tab
            Document doc = new Document(MyDir + "Absolute position tab.docx");

            // An AbsolutePositionTab is a child node of a paragraph
            // AbsolutePositionTabs get picked up when looking for nodes of the SpecialChar type
            Paragraph para = doc.FirstSection.Body.FirstParagraph;
            AbsolutePositionTab absPositionTab = (AbsolutePositionTab)para.GetChild(NodeType.SpecialChar, 0, true);

            // This implementation of the DocumentVisitor pattern converts the document to plain text
            DocToTxtWriter myDocToTxtWriter = new DocToTxtWriter();

            // We can run the DocumentVisitor over the whole first paragraph
            para.Accept(myDocToTxtWriter);

            // A tab character is placed where the AbsolutePositionTab was found
            Assert.AreEqual("Before AbsolutePositionTab\tAfter AbsolutePositionTab", myDocToTxtWriter.GetText());

            // An AbsolutePositionTab can accept a DocumentVisitor by itself too
            myDocToTxtWriter = new DocToTxtWriter();
            absPositionTab.Accept(myDocToTxtWriter);

            Assert.AreEqual("\t", myDocToTxtWriter.GetText());
        }

        /// <summary>
        /// Visitor implementation that simply collects the Runs and AbsolutePositionTabs of a document as plain text. 
        /// </summary>
        public class DocToTxtWriter : DocumentVisitor
        {
            public DocToTxtWriter()
            {
                mBuilder = new StringBuilder();
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
                // We'll treat the AbsolutePositionTab as a regular tab in this case
                mBuilder.Append("\t");
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Adds text to the current output. Honors the enabled/disabled output flag.
            /// </summary>
            private void AppendText(string text)
            {
                mBuilder.Append(text);
            }

            /// <summary>
            /// Gets the plain text of the document that was accumulated by the visitor.
            /// </summary>
            public string GetText()
            {
                return mBuilder.ToString();
            }

            private readonly StringBuilder mBuilder;
        }
        //ExEnd
    }
}