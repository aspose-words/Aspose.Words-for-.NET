// Copyright (c) 2001-2025 Aspose Pty Ltd. All Rights Reserved.
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
            //ExFor:LayoutCollector
            //ExFor:LayoutCollector.#ctor(Document)
            //ExFor:LayoutCollector.Clear
            //ExFor:LayoutCollector.Document
            //ExFor:LayoutCollector.GetEndPageIndex(Node)
            //ExFor:LayoutCollector.GetEntity(Node)
            //ExFor:LayoutCollector.GetNumPagesSpanned(Node)
            //ExFor:LayoutCollector.GetStartPageIndex(Node)
            //ExFor:LayoutEnumerator.Current
            //ExSummary:Shows how to see the the ranges of pages that a node spans.
            Document doc = new Document();
            LayoutCollector layoutCollector = new LayoutCollector(doc);

            // Call the "GetNumPagesSpanned" method to count how many pages the content of our document spans.
            // Since the document is empty, that number of pages is currently zero.
            Assert.That(layoutCollector.Document, Is.EqualTo(doc));
            Assert.That(layoutCollector.GetNumPagesSpanned(doc), Is.EqualTo(0));

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
            Assert.That(layoutCollector.GetNumPagesSpanned(doc), Is.EqualTo(0));

            layoutCollector.Clear();
            doc.UpdatePageLayout();

            Assert.That(layoutCollector.GetNumPagesSpanned(doc), Is.EqualTo(5));

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

            Assert.That(layoutEnumerator.Type, Is.EqualTo(LayoutEntityType.Page));

            // The LayoutEnumerator can traverse the collection of layout entities like a tree.
            // We can also apply it to any node's corresponding layout entity.
            layoutEnumerator.Current = layoutCollector.GetEntity(doc.GetChild(NodeType.Paragraph, 1, true));

            Assert.That(layoutEnumerator.Type, Is.EqualTo(LayoutEntityType.Span));
            Assert.That(layoutEnumerator.Text, Is.EqualTo("¶"));
            //ExEnd
        }

        //ExStart
        //ExFor:LayoutEntityType
        //ExFor:LayoutEnumerator
        //ExFor:LayoutEnumerator.#ctor(Document)
        //ExFor:LayoutEnumerator.Document
        //ExFor:LayoutEnumerator.Kind
        //ExFor:LayoutEnumerator.MoveFirstChild
        //ExFor:LayoutEnumerator.MoveLastChild
        //ExFor:LayoutEnumerator.MoveNext
        //ExFor:LayoutEnumerator.MoveNextLogical
        //ExFor:LayoutEnumerator.MoveParent
        //ExFor:LayoutEnumerator.MoveParent(LayoutEntityType)
        //ExFor:LayoutEnumerator.MovePrevious
        //ExFor:LayoutEnumerator.MovePreviousLogical
        //ExFor:LayoutEnumerator.PageIndex
        //ExFor:LayoutEnumerator.Rectangle
        //ExFor:LayoutEnumerator.Reset
        //ExFor:LayoutEnumerator.Text
        //ExFor:LayoutEnumerator.Type
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

            Assert.That(layoutEnumerator.Document, Is.EqualTo(doc));

            layoutEnumerator.MoveParent(LayoutEntityType.Page);

            Assert.That(layoutEnumerator.Type, Is.EqualTo(LayoutEntityType.Page));
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
        //ExFor:PageLayoutCallbackArgs
        //ExFor:PageLayoutCallbackArgs.Event
        //ExFor:PageLayoutCallbackArgs.Document
        //ExFor:PageLayoutCallbackArgs.PageIndex
        //ExFor:PageLayoutEvent
        //ExFor:LayoutOptions.Callback
        //ExSummary:Shows how to track layout changes with a layout callback.
        [Test]//ExSkip
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

        [Test]
        public void RestartPageNumberingInContinuousSection()
        {
            //ExStart
            //ExFor:LayoutOptions.ContinuousSectionPageNumberingRestart
            //ExFor:ContinuousSectionRestart
            //ExSummary:Shows how to control page numbering in a continuous section.
            Document doc = new Document(MyDir + "Continuous section page numbering.docx");

            // By default Aspose.Words behavior matches the Microsoft Word 2019.
            // If you need old Aspose.Words behavior, repetitive Microsoft Word 2016, use 'ContinuousSectionRestart.FromNewPageOnly'.
            // Page numbering restarts only if there is no other content before the section on the page where the section starts,
            // because of that the numbering will reset to 2 from the second page.
            doc.LayoutOptions.ContinuousSectionPageNumberingRestart = ContinuousSectionRestart.FromNewPageOnly;
            doc.UpdatePageLayout();

            doc.Save(ArtifactsDir + "Layout.RestartPageNumberingInContinuousSection.pdf");
            //ExEnd
        }
    }
}
