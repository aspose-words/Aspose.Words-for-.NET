
using System;
using System.IO;
using System.Drawing;

using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Rendering;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class EnumerateLayoutElements
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

            Document doc = new Document(dataDir + "TestFile.EnumerateLayout.docx");

            // This creates an enumerator which is used to "walk" the elements of a rendered document.
            LayoutEnumerator it = new LayoutEnumerator(doc);

            // This sample uses the enumerator to write information about each layout element to the console.
            LayoutInfoWriter.Run(it);

            // This sample adds a border around each layout element and saves each page as a JPEG image to the data directory.
            OutlineLayoutEntitiesRenderer.Run(doc, it, dataDir);

            Console.WriteLine("\nEnumerate layout elements example ran successfully.");
        }
    }

    class LayoutInfoWriter
    {
        public static void Run(LayoutEnumerator it)
        {
            DisplayLayoutElements(it, string.Empty);
        }

        /// <summary>
        /// Enumerates forward through each layout element in the document and prints out details of each element. 
        /// </summary>
        private static void DisplayLayoutElements(LayoutEnumerator it, string padding)
        {
            do
            {
                DisplayEntityInfo(it, padding);

                if (it.MoveFirstChild())
                {
                    // Recurse into this child element.
                    DisplayLayoutElements(it, AddPadding(padding));
                    it.MoveParent();
                }
            } while (it.MoveNext());
        }

        /// <summary>
        /// Displays information about the current layout entity to the console.
        /// </summary>
        private static void DisplayEntityInfo(LayoutEnumerator it, string padding)
        {
            Console.Write(padding + it.Type + " - " + it.Kind);

            if (it.Type == LayoutEntityType.Span)
                Console.Write(" - " + it.Text);

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

    class OutlineLayoutEntitiesRenderer
    {
        public static void Run(Document doc, LayoutEnumerator it, string folderPath)
        {
            // Make sure the enumerator is at the beginning of the document.
            it.Reset();

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
                        AddBoundingBoxToElementsOnPage(it, g);

                        // Move the enumerator to the next page if there is one.
                        it.MoveNext();

                        img.Save(folderPath + string.Format("TestFile Page {0} Out.png", pageIndex + 1));
                    }
                }
            }
        }

        /// <summary>
        /// Adds a colored border around each layout element on the page.
        /// </summary>
        private static void AddBoundingBoxToElementsOnPage(LayoutEnumerator it, Graphics g)
        {
            do
            {
                // This time instead of MoveFirstChild and MoveNext, we use MoveLastChild and MovePrevious to enumerate from last to first.
                // Enumeration is done backward so the lines of child entities are drawn first and don't overlap the lines of the parent.
                if (it.MoveLastChild())
                {
                    AddBoundingBoxToElementsOnPage(it, g);
                    it.MoveParent();
                }

                // Convert the rectangle representing the position of the layout entity on the page from points to pixels.
                RectangleF rectF = it.Rectangle;
                Rectangle rect = new Rectangle(PointToPixel(rectF.Left, g.DpiX), PointToPixel(rectF.Top, g.DpiY),
                    PointToPixel(rectF.Width, g.DpiX), PointToPixel(rectF.Height, g.DpiY));

                // Draw a line around the layout entity on the page.
                g.DrawRectangle(GetColoredPenFromType(it.Type), rect);

                // Stop after all elements on the page have been procesed.
                if (it.Type == LayoutEntityType.Page)
                    return;

            } while (it.MovePrevious());
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
