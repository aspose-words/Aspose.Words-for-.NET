using System;
using System.IO;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Load_Options
{
    public class WorkingWithTxtLoadOptions : DocsExamplesBase
    {
        [Test]
        public void DetectNumberingWithWhitespaces()
        {
            //ExStart:DetectNumberingWithWhitespaces
            //GistId:ddafc3430967fb4f4f70085fa577d01a
            // Create a plaintext document in the form of a string with parts that may be interpreted as lists.
            // Upon loading, the first three lists will always be detected by Aspose.Words,
            // and List objects will be created for them after loading.
            const string textDoc = "Full stop delimiters:\n" +
                                   "1. First list item 1\n" +
                                   "2. First list item 2\n" +
                                   "3. First list item 3\n\n" +
                                   "Right bracket delimiters:\n" +
                                   "1) Second list item 1\n" +
                                   "2) Second list item 2\n" +
                                   "3) Second list item 3\n\n" +
                                   "Bullet delimiters:\n" +
                                   "• Third list item 1\n" +
                                   "• Third list item 2\n" +
                                   "• Third list item 3\n\n" +
                                   "Whitespace delimiters:\n" +
                                   "1 Fourth list item 1\n" +
                                   "2 Fourth list item 2\n" +
                                   "3 Fourth list item 3";

            // The fourth list, with whitespace inbetween the list number and list item contents,
            // will only be detected as a list if "DetectNumberingWithWhitespaces" in a LoadOptions object is set to true,
            // to avoid paragraphs that start with numbers being mistakenly detected as lists.
            TxtLoadOptions loadOptions = new TxtLoadOptions { DetectNumberingWithWhitespaces = true };

            // Load the document while applying LoadOptions as a parameter and verify the result.
            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(textDoc)), loadOptions);

            doc.Save(ArtifactsDir + "WorkingWithTxtLoadOptions.DetectNumberingWithWhitespaces.docx");
            //ExEnd:DetectNumberingWithWhitespaces
        }

        [Test]
        public void HandleSpacesOptions()
        {
            //ExStart:HandleSpacesOptions
            //GistId:ddafc3430967fb4f4f70085fa577d01a
            const string textDoc = "      Line 1 \n" +
                                   "    Line 2   \n" +
                                   " Line 3       ";

            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                LeadingSpacesOptions = TxtLeadingSpacesOptions.Trim,
                TrailingSpacesOptions = TxtTrailingSpacesOptions.Trim
            };

            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(textDoc)), loadOptions);

            doc.Save(ArtifactsDir + "WorkingWithTxtLoadOptions.HandleSpacesOptions.docx");
            //ExEnd:HandleSpacesOptions
        }

        [Test]
        public void DocumentTextDirection()
        {
            //ExStart:DocumentTextDirection
            //GistId:ddafc3430967fb4f4f70085fa577d01a
            TxtLoadOptions loadOptions = new TxtLoadOptions { DocumentDirection = DocumentDirection.Auto };

            Document doc = new Document(MyDir + "Hebrew text.txt", loadOptions);

            Paragraph paragraph = doc.FirstSection.Body.FirstParagraph;
            Console.WriteLine(paragraph.ParagraphFormat.Bidi);

            doc.Save(ArtifactsDir + "WorkingWithTxtLoadOptions.DocumentTextDirection.docx");
            //ExEnd:DocumentTextDirection
        }

        [Test]
        public void ExportHeadersFootersMode()
        {
            //ExStart:ExportHeadersFootersMode
            //GistId:ddafc3430967fb4f4f70085fa577d01a
            Document doc = new Document();

            // Insert even and primary headers/footers into the document.
            // The primary header/footers will override the even headers/footers.
            doc.FirstSection.HeadersFooters.Add(new HeaderFooter(doc, HeaderFooterType.HeaderEven));
            doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderEven].AppendParagraph("Even header");
            doc.FirstSection.HeadersFooters.Add(new HeaderFooter(doc, HeaderFooterType.FooterEven));
            doc.FirstSection.HeadersFooters[HeaderFooterType.FooterEven].AppendParagraph("Even footer");
            doc.FirstSection.HeadersFooters.Add(new HeaderFooter(doc, HeaderFooterType.HeaderPrimary));
            doc.FirstSection.HeadersFooters[HeaderFooterType.HeaderPrimary].AppendParagraph("Primary header");
            doc.FirstSection.HeadersFooters.Add(new HeaderFooter(doc, HeaderFooterType.FooterPrimary));
            doc.FirstSection.HeadersFooters[HeaderFooterType.FooterPrimary].AppendParagraph("Primary footer");

            // Insert pages to display these headers and footers.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Page 1");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Writeln("Page 2");
            builder.InsertBreak(BreakType.PageBreak);
            builder.Write("Page 3");

            TxtSaveOptions options = new TxtSaveOptions();
            options.SaveFormat = SaveFormat.Text;

            // All headers and footers are placed at the very end of the output document.
            options.ExportHeadersFootersMode = TxtExportHeadersFootersMode.AllAtEnd;
            doc.Save(ArtifactsDir + "WorkingWithTxtLoadOptions.HeadersFootersMode.AllAtEnd.txt", options);

            // Only primary headers and footers are exported at the beginning and end of each section.
            options.ExportHeadersFootersMode = TxtExportHeadersFootersMode.PrimaryOnly;
            doc.Save(ArtifactsDir + "WorkingWithTxtLoadOptions.HeadersFootersMode.PrimaryOnly.txt", options);

            // No headers and footers are exported.
            options.ExportHeadersFootersMode = TxtExportHeadersFootersMode.None;
            doc.Save(ArtifactsDir + "WorkingWithTxtLoadOptions.HeadersFootersMode.None.txt", options);
            //ExEnd:ExportHeadersFootersMode
        }
    }
}
