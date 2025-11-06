using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Save_Options
{
    public class WorkingWithTxtSaveOptions : DocsExamplesBase
    {
        [Test]
        public void AddBidiMarks()
        {
            //ExStart:AddBidiMarks
            //GistId:ddafc3430967fb4f4f70085fa577d01a
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello world!");
            builder.ParagraphFormat.Bidi = true;
            builder.Writeln("שלום עולם!");
            builder.Writeln("مرحبا بالعالم!");

            TxtSaveOptions saveOptions = new TxtSaveOptions { AddBidiMarks = true };

            doc.Save(ArtifactsDir + "WorkingWithTxtSaveOptions.AddBidiMarks.txt", saveOptions);
            //ExEnd:AddBidiMarks
        }

        [Test]
        public void UseTabForListIndentation()
        {
            //ExStart:UseTabForListIndentation
            //GistId:ddafc3430967fb4f4f70085fa577d01a
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a list with three levels of indentation.
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("Item 1");
            builder.ListFormat.ListIndent();
            builder.Writeln("Item 2");
            builder.ListFormat.ListIndent(); 
            builder.Write("Item 3");

            TxtSaveOptions saveOptions = new TxtSaveOptions();
            saveOptions.ListIndentation.Count = 1;
            saveOptions.ListIndentation.Character = '\t';

            doc.Save(ArtifactsDir + "WorkingWithTxtSaveOptions.UseTabForListIndentation.txt", saveOptions);
            //ExEnd:UseTabForListIndentation
        }

        [Test]
        public void UseSpaceForListIndentation()
        {
            //ExStart:UseSpaceForListIndentation
            //GistId:ddafc3430967fb4f4f70085fa577d01a
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Create a list with three levels of indentation.
            builder.ListFormat.ApplyNumberDefault();
            builder.Writeln("Item 1");
            builder.ListFormat.ListIndent();
            builder.Writeln("Item 2");
            builder.ListFormat.ListIndent(); 
            builder.Write("Item 3");

            TxtSaveOptions saveOptions = new TxtSaveOptions();
            saveOptions.ListIndentation.Count = 3;
            saveOptions.ListIndentation.Character = ' ';

            doc.Save(ArtifactsDir + "WorkingWithTxtSaveOptions.UseSpaceForListIndentation.txt", saveOptions);
            //ExEnd:UseSpaceForListIndentation
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