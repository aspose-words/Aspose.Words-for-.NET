#if NET462
using System;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Rendering;
using NUnit.Framework;

namespace DocsExamples.Complex_examples_and_helpers
{
    public class EnumerateLayoutElements : DocsExamplesBase
    {
        [Test]
        public void GetLayoutElements()
        {
            Document doc = new Document(MyDir + "Document layout.docx");

            // Enumerator which is used to "walk" the elements of a rendered document.
            LayoutEnumerator layoutEnumerator = new LayoutEnumerator(doc);

            // Use the enumerator to write information about each layout element to the console.
            LayoutInfoWriter.Run(layoutEnumerator);

            // Adds a border around each layout element and saves each page as a JPEG image to the data directory.
            OutlineLayoutEntitiesRenderer.Run(doc, layoutEnumerator, ArtifactsDir);
        }
    }

    internal class LayoutInfoWriter
    {
        public static void Run(LayoutEnumerator layoutEnumerator)
        {
            DisplayLayoutElements(layoutEnumerator, string.Empty);
        }

        /// <summary>
        /// Enumerates forward through each layout element in the document and prints out details of each element. 
        /// </summary>
        private static void DisplayLayoutElements(LayoutEnumerator layoutEnumerator, string padding)
        {
            do
            {
                DisplayEntityInfo(layoutEnumerator, padding);

                if (layoutEnumerator.MoveFirstChild())
                {
                    // Recurse into this child element.
                    DisplayLayoutElements(layoutEnumerator, AddPadding(padding));
                    layoutEnumerator.MoveParent();
                }
            } while (layoutEnumerator.MoveNext());
        }

        /// <summary>
        /// Displays information about the current layout entity to the console.
        /// </summary>
        private static void DisplayEntityInfo(LayoutEnumerator layoutEnumerator, string padding)
        {
            Console.Write(padding + layoutEnumerator.Type + " - " + layoutEnumerator.Kind);

            if (layoutEnumerator.Type == LayoutEntityType.Span)
                Console.Write(" - " + layoutEnumerator.Text);

            Console.WriteLine();
        }

        /// <summary>
        /// Returns a string of spaces for padding purposes.
        /// </summary>
        private static string AddPadding(string padding)
        {
            return padding + new string(' ', 4);
        }
    }

    internal class OutlineLayoutEntitiesRenderer
    {
        public static void Run(Document doc, LayoutEnumerator layoutEnumerator, string folderPath)
        {
            // Make sure the enumerator is at the beginning of the document.
            layoutEnumerator.Reset();

            for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
            {
                // Use the document class to find information about the current page.
                PageInfo pageInfo = doc.GetPageInfo(pageIndex);

                const float resolution = 150.0f;
                Size pageSize = pageInfo.GetSizeInPixels(1.0f, resolution);

                using (Bitmap img = new Bitmap(pageSize.Width, pageSize.Height))
                {
                    img.SetResolution(resolution, resolution);

                    using (Graphics g = Graphics.FromImage(img))
                    {
                        // Make the background white.
                        g.Clear(Color.White);

                        // Render the page to the graphics.
                        doc.RenderToScale(pageIndex, g, 0.0f, 0.0f, 1.0f);

                        // Add an outline around each element on the page using the graphics object.
                        AddBoundingBoxToElementsOnPage(layoutEnumerator, g);

                        // Move the enumerator to the next page if there is one.
                        layoutEnumerator.MoveNext();

                        img.Save(folderPath + $"EnumerateLayoutElements.Page_{pageIndex + 1}.png");
                    }
                }
            }
        }

        /// <summary>
        /// Adds a colored border around each layout element on the page.
        /// </summary>
        private static void AddBoundingBoxToElementsOnPage(LayoutEnumerator layoutEnumerator, Graphics g)
        {
            do
            {
                // Use MoveLastChild and MovePrevious to enumerate from last to the first enumeration is done backward,
                // so the lines of child entities are drawn first and don't overlap the parent's lines.
                if (layoutEnumerator.MoveLastChild())
                {
                    AddBoundingBoxToElementsOnPage(layoutEnumerator, g);
                    layoutEnumerator.MoveParent();
                }

                // Convert the rectangle representing the position of the layout entity on the page from points to pixels.
                RectangleF rectF = layoutEnumerator.Rectangle;
                Rectangle rect = new Rectangle(PointToPixel(rectF.Left, g.DpiX), PointToPixel(rectF.Top, g.DpiY),
                    PointToPixel(rectF.Width, g.DpiX), PointToPixel(rectF.Height, g.DpiY));

                // Draw a line around the layout entity on the page.
                g.DrawRectangle(GetColoredPenFromType(layoutEnumerator.Type), rect);

                // Stop after all elements on the page have been processed.
                if (layoutEnumerator.Type == LayoutEntityType.Page)
                    return;
            } while (layoutEnumerator.MovePrevious());
        }

        /// <summary>
        /// Returns a different colored pen for each entity type.
        /// </summary>
        private static Pen GetColoredPenFromType(LayoutEntityType type)
        {
            switch (type)
            {
                case LayoutEntityType.Cell:
                    return Pens.Purple;
                case LayoutEntityType.Column:
                    return Pens.Green;
                case LayoutEntityType.Comment:
                    return Pens.LightBlue;
                case LayoutEntityType.Endnote:
                    return Pens.DarkRed;
                case LayoutEntityType.Footnote:
                    return Pens.DarkBlue;
                case LayoutEntityType.HeaderFooter:
                    return Pens.DarkGreen;
                case LayoutEntityType.Line:
                    return Pens.Blue;
                case LayoutEntityType.NoteSeparator:
                    return Pens.LightGreen;
                case LayoutEntityType.Page:
                    return Pens.Red;
                case LayoutEntityType.Row:
                    return Pens.Orange;
                case LayoutEntityType.Span:
                    return Pens.Red;
                case LayoutEntityType.TextBox:
                    return Pens.Yellow;
                default:
                    return Pens.Red;
            }
        }

        /// <summary>
        /// Converts a value in points to pixels.
        /// </summary>
        private static int PointToPixel(float value, double resolution)
        {
            return Convert.ToInt32(ConvertUtil.PointToPixel(value, resolution));
        }
    }
}
#endif