using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Loading_and_Saving
{
    class SpecifyMarkdownSaveOptions
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            SaveAsMD(dataDir);
            SpecifySaveOptionsAndSaveAsMD(dataDir);
            SupportedMarkdownFeatures();
            SetImagesFolder(dataDir);
        }

        private static void SaveAsMD(string dataDir)
        {
            // ExStart:SaveAsMD
            // Load the document from disk.
            Document doc = new Document(dataDir + "Test.docx");

            // Save the document to Markdown format.
            doc.Save(dataDir + "SaveDocx2Markdown.md");
            // ExEnd:SaveAsMD
        }

        private static void SpecifySaveOptionsAndSaveAsMD(string dataDir)
        {
            DocumentBuilder builder = new DocumentBuilder();
            builder.Writeln("Some text!");

            // specify MarkDownSaveOptions
            MarkdownSaveOptions saveOptions = (MarkdownSaveOptions)SaveOptions.CreateSaveOptions(SaveFormat.Markdown);
            
            builder.Document.Save(dataDir + "TestDocument.md", saveOptions);
        }

        private static void SupportedMarkdownFeatures()
        {
            // ExStart:SupportedMarkdownFeatures
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Specify the "Heading 1" style for the paragraph.
            builder.InsertParagraph();
            builder.ParagraphFormat.StyleName = "Heading 1";
            builder.Write("Heading 1");

            // Specify the Italic emphasis for the paragraph.
            builder.InsertParagraph();
            // Reset styles from the previous paragraph to not combine styles between paragraphs.
            builder.ParagraphFormat.StyleName = "Normal";
            builder.Font.Italic = true;
            builder.Write("Italic Text");
            // Reset styles from the previous paragraph to not combine styles between paragraphs.
            builder.Italic = false;

            // Specify a Hyperlink for the desired text.
            builder.InsertParagraph();
            builder.InsertHyperlink("Aspose", "https://www.aspose.com", false);
            builder.Write("Aspose");
            
            // Save your document as a Markdown file.
            doc.Save("example.md");
            // ExEnd:SupportedMarkdownFeatures
        }

        private static void SetImagesFolder(string dataDir)
        {
            // ExStart:SetImagesFolder
            // Load the document from disk.
            Document doc = new Document(dataDir + "Test.docx");

            MarkdownSaveOptions so = new MarkdownSaveOptions();
            so.ImagesFolder = dataDir + "\\Images";
            
            using (MemoryStream stream = new MemoryStream())
                doc.Save(stream, so);
            // ExEnd:SetImagesFolder
        }
    }
}
