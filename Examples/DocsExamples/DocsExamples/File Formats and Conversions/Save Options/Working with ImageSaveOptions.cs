using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace DocsExamples.File_Formats_and_Conversions.Save_Options
{
    public class WorkingWithImageSaveOptions : DocsExamplesBase
    {
        [Test]
        public void ExposeThresholdControl()
        {
            //ExStart:ExposeThresholdControl
            //GistId:b20a0ec0e1ff0556aa20d12f486e1963
            Document doc = new Document(MyDir + "Rendering.docx");

            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                TiffCompression = TiffCompression.Ccitt3,
                ImageColorMode = ImageColorMode.Grayscale,
                TiffBinarizationMethod = ImageBinarizationMethod.FloydSteinbergDithering,
                ThresholdForFloydSteinbergDithering = 254
            };

            doc.Save(ArtifactsDir + "WorkingWithImageSaveOptions.ExposeThresholdControl.tiff", saveOptions);
            //ExEnd:ExposeThresholdControl
        }

        [Test]
        public void GetTiffPageRange()
        {
            //ExStart:GetTiffPageRange
            //GistId:b20a0ec0e1ff0556aa20d12f486e1963
            Document doc = new Document(MyDir + "Rendering.docx");
            //ExStart:SaveAsTiff
            //GistId:b20a0ec0e1ff0556aa20d12f486e1963
            doc.Save(ArtifactsDir + "WorkingWithImageSaveOptions.MultipageTiff.tiff");
            //ExEnd:SaveAsTiff

            //ExStart:SaveAsTIFFUsingImageSaveOptions
            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Tiff)
            {
                PageSet = new PageSet(new PageRange(0, 1)), TiffCompression = TiffCompression.Ccitt4, Resolution = 160
            };

            doc.Save(ArtifactsDir + "WorkingWithImageSaveOptions.GetTiffPageRange.tiff", saveOptions);
            //ExEnd:SaveAsTIFFUsingImageSaveOptions
            //ExEnd:GetTiffPageRange
        }

        [Test]
        public void Format1BppIndexed()
        {
            //ExStart:Format1BppIndexed
            //GistId:83e5c469d0e72b5114fb8a05a1d01977
            Document doc = new Document(MyDir + "Rendering.docx");

            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                PageSet = new PageSet(1),
                ImageColorMode = ImageColorMode.BlackAndWhite,
                PixelFormat = ImagePixelFormat.Format1bppIndexed
            };

            doc.Save(ArtifactsDir + "WorkingWithImageSaveOptions.Format1BppIndexed.Png", saveOptions);
            //ExEnd:Format1BppIndexed
        }

        [Test]
        public void GetJpegPageRange()
        {
            //ExStart:GetJpegPageRange
            //GistId:ebbb90d74ef57db456685052a18f8e86
            Document doc = new Document(MyDir + "Rendering.docx");

            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Jpeg);

            // Set the "PageSet" to "0" to convert only the first page of a document.
            options.PageSet = new PageSet(0);

            // Change the image's brightness and contrast.
            // Both are on a 0-1 scale and are at 0.5 by default.
            options.ImageBrightness = 0.3f;
            options.ImageContrast = 0.7f;

            // Change the horizontal resolution.
            // The default value for these properties is 96.0, for a resolution of 96dpi.
            options.HorizontalResolution = 72f;

            doc.Save(ArtifactsDir + "WorkingWithImageSaveOptions.GetJpegPageRange.jpeg", options);
            //ExEnd:GetJpegPageRange
        }

        [Test]
        //ExStart:PageSavingCallback
        public static void PageSavingCallback()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(SaveFormat.Png)
            {
                PageSet = new PageSet(new PageRange(0, doc.PageCount - 1)),
                PageSavingCallback = new HandlePageSavingCallback()
            };

            doc.Save(ArtifactsDir + "WorkingWithImageSaveOptions.PageSavingCallback.png", imageSaveOptions);
        }

        private class HandlePageSavingCallback : IPageSavingCallback
        {
            public void PageSaving(PageSavingArgs args)
            {
                args.PageFileName = string.Format(ArtifactsDir + "Page_{0}.png", args.PageIndex);
            }
        }
        //ExEnd:PageSavingCallback
    }
}