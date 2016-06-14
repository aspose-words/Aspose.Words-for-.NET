using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;
using Aspose.Words.Drawing;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Images
{
    class AddImageToEachPage
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithImages();
            string fileName = "TestFile.doc";
            // This a document that we want to add an image and custom text for each page without using the header or footer.
            Document doc = new Document(dataDir + fileName);

            // Create and attach collector before the document before page layout is built.
            LayoutCollector layoutCollector = new LayoutCollector(doc);

            // Images in a document are added to paragraphs, so to add an image to every page we need to find at any paragraph 
            // belonging to each page.
            IEnumerator enumerator = doc.SelectNodes("//Body/Paragraph").GetEnumerator();

            // Loop through each document page.
            for (int page = 1; page <= doc.PageCount; page++)
            {
                while (enumerator.MoveNext())
                {
                    // Check if the current paragraph belongs to the target page.
                    Paragraph paragraph = (Paragraph)enumerator.Current;
                    if (layoutCollector.GetStartPageIndex(paragraph) == page)
                    {
                        AddImageToPage(paragraph, page, dataDir);
                        break;
                    }
                }
            }

            // Call UpdatePageLayout() method if file is to be saved as PDF or image format
            doc.UpdatePageLayout();

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            doc.Save(dataDir);

            Console.WriteLine("\nInserted images on each page of the document successfully.\nFile saved at " + dataDir);
        }

        /// <summary>
        /// Adds an image to a page using the supplied paragraph.
        /// </summary>
        /// <param name="para">The paragraph to an an image to.</param>
        /// <param name="page">The page number the paragraph appears on.</param>
        public static void AddImageToPage(Paragraph para, int page, String dataDir)
        {
            Document doc = (Document)para.Document;

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.MoveTo(para);

            // Add a logo to the top left of the page. The image is placed infront of all other text.
            Shape shape = builder.InsertImage(dataDir + "Aspose Logo.png", RelativeHorizontalPosition.Page, 60,
                RelativeVerticalPosition.Page, 60, -1, -1, WrapType.None);

            // Add a textbox next to the image which contains some text consisting of the page number. 
            Shape textBox = new Shape(doc, ShapeType.TextBox);

            // We want a floating shape relative to the page.
            textBox.WrapType = WrapType.None;
            textBox.RelativeHorizontalPosition = RelativeHorizontalPosition.Page;
            textBox.RelativeVerticalPosition = RelativeVerticalPosition.Page;

            // Set the textbox position.
            textBox.Height = 30;
            textBox.Width = 200;
            textBox.Left = 150;
            textBox.Top = 80;

            // Add the textbox and set text.
            textBox.AppendChild(new Paragraph(doc));
            builder.InsertNode(textBox);
            builder.MoveTo(textBox.FirstChild);
            builder.Writeln("This is a custom note for page " + page);
        }
    }
}
