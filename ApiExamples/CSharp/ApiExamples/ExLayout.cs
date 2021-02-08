// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
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
            //ExSummary:Shows how to see the the ranges of pages that a node spans.
            Document doc = new Document();
            LayoutCollector layoutCollector = new LayoutCollector(doc);
            
            // Call the "GetNumPagesSpanned" method to count how many pages the content of our document spans.
            // Since the document is empty, that number of pages is currently zero.
            Assert.AreEqual(doc, layoutCollector.Document);
            Assert.AreEqual(0, layoutCollector.GetNumPagesSpanned(doc));

            // Populate the document with 5 pages of content.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("Section 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.InsertBreak(BreakType.PageBreak);
            builder.InsertBreak(BreakType.SectionBreakEvenPage);
            builder.Write("Section 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.InsertBreak(BreakType.PageBreak);

            // Before the layout collector, we need to call the "UpdatePageLayout" method to give us
            // an accurate figure for any layout-related metric, such as the page count.
            Assert.AreEqual(0, layoutCollector.GetNumPagesSpanned(doc));

            layoutCollector.Clear();
            doc.UpdatePageLayout();

            Assert.AreEqual(5, layoutCollector.GetNumPagesSpanned(doc));

            // We can see the numbers of the start and end pages of any node and their overall page spans.
            NodeCollection nodes = doc.GetChildNodes(NodeType.Any, true);
            foreach (Node node in nodes)
            {
                Console.WriteLine($"->  NodeType.{node.NodeType}: ");
                Console.WriteLine(
                    $"\tStarts on page {layoutCollector.GetStartPageIndex(node)}, ends on page {layoutCollector.GetEndPageIndex(node)}," +
                    $" spanning {layoutCollector.GetNumPagesSpanned(node)} pages.");
            }

            // We can iterate over the layout entities using a LayoutEnumerator.
            LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);

            Assert.AreEqual(LayoutEntityType.Page, layoutEnumerator.Type);

            // The LayoutEnumerator can traverse the collection of layout entities like a tree.
            // We can also apply it to any node's corresponding layout entity.
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
            // Open a document that contains a variety of layout entities.
            // Layout entities are pages, cells, rows, lines, and other objects included in the LayoutEntityType enum.
            // Each layout entity has a rectangular space that it occupies in the document body.
            Document doc = new Document(MyDir + "Layout entities.docx");

            // Create an enumerator that can traverse these entities like a tree.
            LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);

            Assert.AreEqual(doc, layoutEnumerator.Document);

            layoutEnumerator.MoveParent(LayoutEntityType.Page);

            Assert.AreEqual(LayoutEntityType.Page, layoutEnumerator.Type);
            Assert.Throws<InvalidOperationException>(() => Console.WriteLine(layoutEnumerator.Text));

            // We can call this method to make sure that the enumerator will be at the first layout entity.
            layoutEnumerator.Reset();

            // There are two orders that determine how the layout enumerator continues traversing layout entities
            // when it encounters entities that span across multiple pages.
            // 1 -  In visual order:
            // When moving through an entity's children that span multiple pages,
            // page layout takes precedence, and we move to other child elements on this page and avoid the ones on the next.
            Console.WriteLine("Traversing from first to last, elements between pages separated:");
            TraverseLayoutForward(layoutEnumerator, 1);

            // Our enumerator is now at the end of the collection. We can traverse the layout entities backwards to go back to the beginning.
            Console.WriteLine("Traversing from last to first, elements between pages separated:");
            TraverseLayoutBackward(layoutEnumerator, 1);

            // 2 -  In logical order:
            // When moving through an entity's children that span multiple pages,
            // the enumerator will move between pages to traverse all the child entities.
            Console.WriteLine("Traversing from first to last, elements between pages mixed:");
            TraverseLayoutForwardLogical(layoutEnumerator, 1);

            Console.WriteLine("Traversing from last to first, elements between pages mixed:");
            TraverseLayoutBackwardLogical(layoutEnumerator, 1);
        }

        /// <summary>
        /// Enumerate through layoutEnumerator's layout entity collection front-to-back,
        /// in a depth-first manner, and in the "Visual" order.
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
        /// Enumerate through layoutEnumerator's layout entity collection back-to-front,
        /// in a depth-first manner, and in the "Visual" order.
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
        /// Enumerate through layoutEnumerator's layout entity collection front-to-back,
        /// in a depth-first manner, and in the "Logical" order.
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
        /// Enumerate through layoutEnumerator's layout entity collection back-to-front,
        /// in a depth-first manner, and in the "Logical" order.
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
        /// Print information about layoutEnumerator's current entity to the console, while indenting the text with tab characters
        /// based on its depth relative to the root node that we provided in the constructor LayoutEnumerator instance.
        /// The rectangle that we process at the end represents the area and location that the entity takes up in the document.
        /// </summary>
        private static void PrintCurrentEntity(LayoutEnumerator layoutEnumerator, int indent)
        {
            string tabs = new string('\t', indent);

            Console.WriteLine(layoutEnumerator.Kind == string.Empty
                ? $"{tabs}-> Entity type: {layoutEnumerator.Type}"
                : $"{tabs}-> Entity type & kind: {layoutEnumerator.Type}, {layoutEnumerator.Kind}");

            // Only spans can contain text.
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
        //ExSummary:Shows how to track layout changes with a layout callback.
        [Test]
        public void PageLayoutCallback()
        {
            Document doc = new Document();
            doc.BuiltInDocumentProperties.Title = "My Document";

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            doc.LayoutOptions.Callback = new RenderPageLayoutCallback();
            doc.UpdatePageLayout();

            doc.Save(ArtifactsDir + "Layout.PageLayoutCallback.pdf");
        }

        /// <summary>
        /// Notifies us when we save the document to a fixed page format
        /// and renders a page that we perform a page reflow on to an image in the local file system.
        /// </summary>
        private class RenderPageLayoutCallback : IPageLayoutCallback
        {
            public void Notify(PageLayoutCallbackArgs a)
            {
                switch (a.Event)
                {
                    case PageLayoutEvent.PartReflowFinished:
                        NotifyPartFinished(a);
                        break;
                    case PageLayoutEvent.ConversionFinished:
                        NotifyConversionFinished(a);
                        break;
                }
            }

            private void NotifyPartFinished(PageLayoutCallbackArgs a)
            {
                Console.WriteLine($"Part at page {a.PageIndex + 1} reflow.");
                RenderPage(a, a.PageIndex);
            }

            private void NotifyConversionFinished(PageLayoutCallbackArgs a)
            {
                Console.WriteLine($"Document \"{a.Document.BuiltInDocumentProperties.Title}\" converted to page format.");
            }

            private void RenderPage(PageLayoutCallbackArgs a, int pageIndex)
            {
                ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png) { PageSet = new PageSet(pageIndex) };

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
