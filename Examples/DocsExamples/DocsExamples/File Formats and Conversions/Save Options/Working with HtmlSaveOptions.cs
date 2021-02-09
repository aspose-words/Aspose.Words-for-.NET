using System.IO;
using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Save_Options
{
    public class WorkingWithHtmlSaveOptions : DocsExamplesBase
    {
        [Test]
        public void ExportRoundtripInformation()
        {
            //ExStart:ExportRoundtripInformation
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions { ExportRoundtripInformation = true };

            doc.Save(ArtifactsDir + "WorkingWithHtmlSaveOptions.ExportRoundtripInformation.html", saveOptions);
            //ExEnd:ExportRoundtripInformation
        }

        [Test]
        public void ExportFontsAsBase64()
        {
            //ExStart:ExportFontsAsBase64
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions { ExportFontsAsBase64 = true };

            doc.Save(ArtifactsDir + "WorkingWithHtmlSaveOptions.ExportFontsAsBase64.html", saveOptions);
            //ExEnd:ExportFontsAsBase64
        }

        [Test]
        public void ExportResources()
        {
            //ExStart:ExportResources
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                CssStyleSheetType = CssStyleSheetType.External,
                ExportFontResources = true,
                ResourceFolder = ArtifactsDir + "Resources",
                ResourceFolderAlias = "http://example.com/resources"
            };

            doc.Save(ArtifactsDir + "WorkingWithHtmlSaveOptions.ExportResources.html", saveOptions);
            //ExEnd:ExportResources
        }

        [Test]
        public void ConvertMetafilesToEmfOrWmf()
        {
            //ExStart:ConvertMetafilesToEmfOrWmf
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Write("Here is an image as is: ");
            builder.InsertHtml(
                @"<img src=""data:image/png;base64,
                    iVBORw0KGgoAAAANSUhEUgAAAAoAAAAKCAYAAACNMs+9AAAABGdBTUEAALGP
                    C/xhBQAAAAlwSFlzAAALEwAACxMBAJqcGAAAAAd0SU1FB9YGARc5KB0XV+IA
                    AAAddEVYdENvbW1lbnQAQ3JlYXRlZCB3aXRoIFRoZSBHSU1Q72QlbgAAAF1J
                    REFUGNO9zL0NglAAxPEfdLTs4BZM4DIO4C7OwQg2JoQ9LE1exdlYvBBeZ7jq
                    ch9//q1uH4TLzw4d6+ErXMMcXuHWxId3KOETnnXXV6MJpcq2MLaI97CER3N0
                    vr4MkhoXe0rZigAAAABJRU5ErkJggg=="" alt=""Red dot"" />");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions { MetafileFormat = HtmlMetafileFormat.EmfOrWmf };

            doc.Save(ArtifactsDir + "WorkingWithHtmlSaveOptions.ConvertMetafilesToEmfOrWmf.html", saveOptions);
            //ExEnd:ConvertMetafilesToEmfOrWmf
        }

        [Test]
        public void ConvertMetafilesToSvg()
        {
            //ExStart:ConvertMetafilesToSvg
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            builder.Write("Here is an SVG image: ");
            builder.InsertHtml(
                @"<svg height='210' width='500'>
                <polygon points='100,10 40,198 190,78 10,78 160,198' 
                    style='fill:lime;stroke:purple;stroke-width:5;fill-rule:evenodd;' />
            </svg> ");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions { MetafileFormat = HtmlMetafileFormat.Svg };

            doc.Save(ArtifactsDir + "WorkingWithHtmlSaveOptions.ConvertMetafilesToSvg.html", saveOptions);
            //ExEnd:ConvertMetafilesToSvg
        }

        [Test]
        public void AddCssClassNamePrefix()
        {
            //ExStart:AddCssClassNamePrefix
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                CssStyleSheetType = CssStyleSheetType.External, CssClassNamePrefix = "pfx_"
            };
            
            doc.Save(ArtifactsDir + "WorkingWithHtmlSaveOptions.AddCssClassNamePrefix.html", saveOptions);
            //ExEnd:AddCssClassNamePrefix
        }

        [Test]
        public void ExportCidUrlsForMhtmlResources()
        {
            //ExStart:ExportCidUrlsForMhtmlResources
            Document doc = new Document(MyDir + "Content-ID.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Mhtml)
            {
                PrettyFormat = true, ExportCidUrlsForMhtmlResources = true
            };

            doc.Save(ArtifactsDir + "WorkingWithHtmlSaveOptions.ExportCidUrlsForMhtmlResources.mhtml", saveOptions);
            //ExEnd:ExportCidUrlsForMhtmlResources
        }

        [Test]
        public void ResolveFontNames()
        {
            //ExStart:ResolveFontNames
            Document doc = new Document(MyDir + "Missing font.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                PrettyFormat = true, ResolveFontNames = true
            };

            doc.Save(ArtifactsDir + "WorkingWithHtmlSaveOptions.ResolveFontNames.html", saveOptions);
            //ExEnd:ResolveFontNames
        }

        [Test]
        public void ExportTextInputFormFieldAsText()
        {
            //ExStart:ExportTextInputFormFieldAsText
            Document doc = new Document(MyDir + "Rendering.docx");

            string imagesDir = Path.Combine(ArtifactsDir, "Images");

            // The folder specified needs to exist and should be empty.
            if (Directory.Exists(imagesDir))
                Directory.Delete(imagesDir, true);

            Directory.CreateDirectory(imagesDir);

            // Set an option to export form fields as plain text, not as HTML input elements.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportTextInputFormFieldAsText = true, ImagesFolder = imagesDir
            };

            doc.Save(ArtifactsDir + "WorkingWithHtmlSaveOptions.ExportTextInputFormFieldAsText.html", saveOptions);
            //ExEnd:ExportTextInputFormFieldAsText
        }
    }
}