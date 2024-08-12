using System.IO;
using System.Text;
using Aspose.Words;
using Aspose.Words.Loading;
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
            //GistId:c0df00d37081f41a7683339fd7ef66c1
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions { ExportRoundtripInformation = true };

            doc.Save(ArtifactsDir + "WorkingWithHtmlSaveOptions.ExportRoundtripInformation.html", saveOptions);
            //ExEnd:ExportRoundtripInformation
        }

        [Test]
        public void ExportFontsAsBase64()
        {
            //ExStart:ExportFontsAsBase64
            //GistId:c0df00d37081f41a7683339fd7ef66c1
            Document doc = new Document(MyDir + "Rendering.docx");

            HtmlSaveOptions saveOptions = new HtmlSaveOptions { ExportFontsAsBase64 = true };

            doc.Save(ArtifactsDir + "WorkingWithHtmlSaveOptions.ExportFontsAsBase64.html", saveOptions);
            //ExEnd:ExportFontsAsBase64
        }

        [Test]
        public void ExportResources()
        {
            //ExStart:ExportResources
            //GistId:c0df00d37081f41a7683339fd7ef66c1
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
        public void ConvertMetafilesToPng()
        {
            //ExStart:ConvertMetafilesToPng
            string html =
                @"<html>
                    <svg xmlns='http://www.w3.org/2000/svg' width='500' height='40' viewBox='0 0 500 40'>
                        <text x='0' y='35' font-family='Verdana' font-size='35'>Hello world!</text>
                    </svg>
                </html>";

            // Use 'ConvertSvgToEmf' to turn back the legacy behavior
            // where all SVG images loaded from an HTML document were converted to EMF.
            // Now SVG images are loaded without conversion
            // if the MS Word version specified in load options supports SVG images natively.
            HtmlLoadOptions loadOptions = new HtmlLoadOptions { ConvertSvgToEmf = true };
            Document doc = new Document(new MemoryStream(Encoding.UTF8.GetBytes(html)), loadOptions);

            HtmlSaveOptions saveOptions = new HtmlSaveOptions { MetafileFormat = HtmlMetafileFormat.Png };

            doc.Save(ArtifactsDir + "WorkingWithHtmlSaveOptions.ConvertMetafilesToPng.html", saveOptions);
            //ExEnd:ConvertMetafilesToPng
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
            //GistId:83e5c469d0e72b5114fb8a05a1d01977
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