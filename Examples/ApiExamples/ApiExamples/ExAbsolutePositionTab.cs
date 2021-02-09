// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
        //ExSummary:Shows how to process absolute position tab characters with a document visitor.
        [Test] //ExSkip
        public void DocumentToTxt()
        {
            Document doc = new Document(MyDir + "Absolute position tab.docx");

            // Extract the text contents of our document by accepting this custom document visitor.
            DocTextExtractor myDocTextExtractor = new DocTextExtractor();
            doc.FirstSection.Body.Accept(myDocTextExtractor);

            // The absolute position tab, which has no equivalent in string form, has been explicitly converted to a tab character.
            Assert.AreEqual("Before AbsolutePositionTab\tAfter AbsolutePositionTab", myDocTextExtractor.GetText());

            // An AbsolutePositionTab can accept a DocumentVisitor by itself too.
            AbsolutePositionTab absPositionTab = (AbsolutePositionTab)doc.FirstSection.Body.FirstParagraph.GetChild(NodeType.SpecialChar, 0, true);

            myDocTextExtractor = new DocTextExtractor();
            absPositionTab.Accept(myDocTextExtractor);

            Assert.AreEqual("\t", myDocTextExtractor.GetText());
        }

        /// <summary>
        /// Collects the text contents of all runs in the visited document. Replaces all absolute tab characters with ordinary tabs.
        /// </summary>
        public class DocTextExtractor : DocumentVisitor
        {
            public DocTextExtractor()
            {
                mBuilder = new StringBuilder();
            }

            /// <summary>
            /// Called when a Run node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitRun(Run run)
            {
                AppendText(run.Text);
                return VisitorAction.Continue;
            }

            /// <summary>
            /// Called when an AbsolutePositionTab node is encountered in the document.
            /// </summary>
            public override VisitorAction VisitAbsolutePositionTab(AbsolutePositionTab tab)
            {
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
            /// Plain text of the document that was accumulated by the visitor.
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