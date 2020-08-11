// Copyright (c) 2001-2020 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Drawing;
using System.IO;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    class ExLayout : ApiExampleBase
    {
        [Test]
        public void LayoutCollector()
        {
            //ExStart
            //ExFor:Layout.LayoutCollector
            //ExFor:Layout.LayoutCollector.#ctor(Document)
            //ExFor:Layout.LayoutCollector.Clear
            //ExFor:Layout.LayoutCollector.Document
            //ExFor:Layout.LayoutCollector.GetEndPageIndex(Node)
            //ExFor:Layout.LayoutCollector.GetEntity(Node)
            //ExFor:Layout.LayoutCollector.GetNumPagesSpanned(Node)
            //ExFor:Layout.LayoutCollector.GetStartPageIndex(Node)
            //ExFor:Layout.LayoutEnumerator.Current
            //ExSummary:Shows how to see the page spans of nodes.
            // Open a blank document and create a DocumentBuilder
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a LayoutCollector object for our document that will have information about the nodes we placed
            LayoutCollector layoutCollector = new LayoutCollector(doc);

            // The document itself is a node that contains everything, which currently spans 0 pages
            Assert.AreEqual(doc, layoutCollector.Document);
            Assert.AreEqual(0, layoutCollector.GetNumPagesSpanned(doc));

            // Populate the document with sections and page breaks
            builder.Write("Section 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.InsertBreak(BreakType.PageBreak);
            doc.AppendChild(new Section(doc));
            doc.LastSection.AppendChild(new Body(doc));
            builder.MoveToDocumentEnd();
            builder.Write("Section 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.InsertBreak(BreakType.PageBreak);

            // The collected layout data won't automatically keep up with the real document contents
            Assert.AreEqual(0, layoutCollector.GetNumPagesSpanned(doc));

            // After we clear the layout collection and update it, the layout entity collection will be populated with up-to-date information about our nodes
            // The page span for the document now shows 5, which is what we would expect after placing 4 page breaks
            layoutCollector.Clear();
            doc.UpdatePageLayout();
            Assert.AreEqual(5, layoutCollector.GetNumPagesSpanned(doc));

            // We can also see the start/end pages of any other node, and their overall page spans
            NodeCollection nodes = doc.GetChildNodes(NodeType.Any, true);
            foreach (Node node in nodes)
            {
                Console.WriteLine($"->  NodeType.{node.NodeType}: ");
                Console.WriteLine(
                    $"\tStarts on page {layoutCollector.GetStartPageIndex(node)}, ends on page {layoutCollector.GetEndPageIndex(node)}," +
                    $" spanning {layoutCollector.GetNumPagesSpanned(node)} pages.");
            }

            // We can iterate over the layout entities using a LayoutEnumerator
            LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);
            Assert.AreEqual(LayoutEntityType.Page, layoutEnumerator.Type);

            // The LayoutEnumerator can traverse the collection of layout entities like a tree
            // We can also point it to any node's corresponding layout entity like this
            layoutEnumerator.Current = layoutCollector.GetEntity(doc.GetChild(NodeType.Paragraph, 1, true));
            Assert.AreEqual(LayoutEntityType.Span, layoutEnumerator.Type);
            Assert.AreEqual("¶", layoutEnumerator.Text);
            //ExEnd
        }

        //ExStart
        //ExFor:Layout.LayoutEntityType
        //ExFor:Layout.LayoutEnumerator
        //ExFor:Layout.LayoutEnumerator.#ctor(Document)
        //ExFor:Layout.LayoutEnumerator.Document
        //ExFor:Layout.LayoutEnumerator.Kind
        //ExFor:Layout.LayoutEnumerator.MoveFirstChild
        //ExFor:Layout.LayoutEnumerator.MoveLastChild
        //ExFor:Layout.LayoutEnumerator.MoveNext
        //ExFor:Layout.LayoutEnumerator.MoveNextLogical
        //ExFor:Layout.LayoutEnumerator.MoveParent
        //ExFor:Layout.LayoutEnumerator.MoveParent(Layout.LayoutEntityType)
        //ExFor:Layout.LayoutEnumerator.MovePrevious
        //ExFor:Layout.LayoutEnumerator.MovePreviousLogical
        //ExFor:Layout.LayoutEnumerator.PageIndex
        //ExFor:Layout.LayoutEnumerator.Rectangle
        //ExFor:Layout.LayoutEnumerator.Reset
        //ExFor:Layout.LayoutEnumerator.Text
        //ExFor:Layout.LayoutEnumerator.Type
        //ExSummary:Shows ways of traversing a document's layout entities.
        [Test] //ExSkip
        public void LayoutEnumerator()
        {
            // Open a document that contains a variety of layout entities
            // Layout entities are pages, cells, rows, lines and other objects included in the LayoutEntityType enum
            // They are defined visually by the rectangular space that they occupy in the document
            Document doc = new Document(MyDir + "Layout entities.docx");

            // Create an enumerator that can traverse these entities like a tree
            LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);
            Assert.AreEqual(doc, layoutEnumerator.Document);

            layoutEnumerator.MoveParent(LayoutEntityType.Page);
            Assert.AreEqual(LayoutEntityType.Page, layoutEnumerator.Type);
            Assert.Throws<InvalidOperationException>(() => Console.WriteLine(layoutEnumerator.Text));

            // We can call this method to make sure that the enumerator points to the very first entity before we go through it forwards
            layoutEnumerator.Reset();

            // "Visual order" means when moving through an entity's children that are broken across pages,
            // page layout takes precedence and we avoid elements in other pages and move to others on the same page
            Console.WriteLine("Traversing from first to last, elements between pages separated:");
            TraverseLayoutForward(layoutEnumerator, 1);

            // Our enumerator is conveniently at the end of the collection for us to go through the collection backwards
            Console.WriteLine("Traversing from last to first, elements between pages separated:");
            TraverseLayoutBackward(layoutEnumerator, 1);

            // "Logical order" means when moving through an entity's children that are broken across pages, 
            // node relationships take precedence
            Console.WriteLine("Traversing from first to last, elements between pages mixed:");
            TraverseLayoutForwardLogical(layoutEnumerator, 1);

            Console.WriteLine("Traversing from last to first, elements between pages mixed:");
            TraverseLayoutBackwardLogical(layoutEnumerator, 1);
        }

        /// <summary>
        /// Enumerate through layoutEnumerator's layout entity collection front-to-back, in a DFS manner, and in a "Visual" order.
        /// </summary>
        private static void TraverseLayoutForward(LayoutEnumerator layoutEnumerator, int depth)
        {
            do
            {
                PrintCurrentEntity(layoutEnumerator, depth);

                if (layoutEnumerator.MoveFirstChild())
                {
                    TraverseLayoutForward(layoutEnumerator, depth + 1);
                    layoutEnumerator.MoveParent();
                }
            } while (layoutEnumerator.MoveNext());
        }

        /// <summary>
        /// Enumerate through layoutEnumerator's layout entity collection back-to-front, in a DFS manner, and in a "Visual" order.
        /// </summary>
        private static void TraverseLayoutBackward(LayoutEnumerator layoutEnumerator, int depth)
        {
            do
            {
                PrintCurrentEntity(layoutEnumerator, depth);

                if (layoutEnumerator.MoveLastChild())
                {
                    TraverseLayoutBackward(layoutEnumerator, depth + 1);
                    layoutEnumerator.MoveParent();
                }
            } while (layoutEnumerator.MovePrevious());
        }

        /// <summary>
        /// Enumerate through layoutEnumerator's layout entity collection front-to-back, in a DFS manner, and in a "Logical" order.
        /// </summary>
        private static void TraverseLayoutForwardLogical(LayoutEnumerator layoutEnumerator, int depth)
        {
            do
            {
                PrintCurrentEntity(layoutEnumerator, depth);

                if (layoutEnumerator.MoveFirstChild())
                {
                    TraverseLayoutForwardLogical(layoutEnumerator, depth + 1);
                    layoutEnumerator.MoveParent();
                }
            } while (layoutEnumerator.MoveNextLogical());
        }

        /// <summary>
        /// Enumerate through layoutEnumerator's layout entity collection back-to-front, in a DFS manner, and in a "Logical" order.
        /// </summary>
        private static void TraverseLayoutBackwardLogical(LayoutEnumerator layoutEnumerator, int depth)
        {
            do
            {
                PrintCurrentEntity(layoutEnumerator, depth);

                if (layoutEnumerator.MoveLastChild())
                {
                    TraverseLayoutBackwardLogical(layoutEnumerator, depth + 1);
                    layoutEnumerator.MoveParent();
                }
            } while (layoutEnumerator.MovePreviousLogical());
        }

        /// <summary>
        /// Print information about layoutEnumerator's current entity to the console, indented by a number of tab characters specified by indent.
        /// The rectangle that we process at the end represents the area and location thereof that the element takes up in the document.
        /// </summary>
        private static void PrintCurrentEntity(LayoutEnumerator layoutEnumerator, int indent)
        {
            string tabs = new string('\t', indent);

            Console.WriteLine(layoutEnumerator.Kind == string.Empty
                ? $"{tabs}-> Entity type: {layoutEnumerator.Type}"
                : $"{tabs}-> Entity type & kind: {layoutEnumerator.Type}, {layoutEnumerator.Kind}");

            // Only spans can contain text
            if (layoutEnumerator.Type == LayoutEntityType.Span)
                Console.WriteLine($"{tabs}   Span contents: \"{layoutEnumerator.Text}\"");

            RectangleF leRect = layoutEnumerator.Rectangle;
            Console.WriteLine($"{tabs}   Rectangle dimensions {leRect.Width}x{leRect.Height}, X={leRect.X} Y={leRect.Y}");
            Console.WriteLine($"{tabs}   Page {layoutEnumerator.PageIndex}");
        }
        //ExEnd

        //ExStart
        //ExFor:IPageLayoutCallback
        //ExFor:IPageLayoutCallback.Notify(PageLayoutCallbackArgs)
        //ExFor:PageLayoutCallbackArgs.Event
        //ExFor:PageLayoutCallbackArgs.Document
        //ExFor:PageLayoutCallbackArgs.PageIndex
        //ExFor:PageLayoutEvent
        //ExSummary:Shows how to track layout/rendering progress with layout callback.
        [Test]
        public void PageLayoutCallback()
        {
            Document doc = new Document(MyDir + "Document.docx");

            doc.LayoutOptions.Callback = new RenderPageLayoutCallback();
            doc.UpdatePageLayout();
        }

        private class RenderPageLayoutCallback : IPageLayoutCallback
        {
            public void Notify(PageLayoutCallbackArgs a)
            {
                switch (a.Event)
                {
                    case PageLayoutEvent.PartReflowFinished:
                        NotifyPartFinished(a);
                        break;
                }
            }

            private void NotifyPartFinished(PageLayoutCallbackArgs a)
            {
                Console.WriteLine($"Part at page {a.PageIndex + 1} reflow");
                RenderPage(a, a.PageIndex);
            }

            private void RenderPage(PageLayoutCallbackArgs a, int pageIndex)
            {
                ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png);
                saveOptions.PageIndex = pageIndex;
                saveOptions.PageCount = 1;

                using (FileStream stream =
                    new FileStream(ArtifactsDir + $@"PageLayoutCallback.page-{pageIndex + 1} {++mNum}.png",
                        FileMode.Create))
                    a.Document.Save(stream, saveOptions);
            }

            private int mNum;
        }
        //ExEnd
    }
}
