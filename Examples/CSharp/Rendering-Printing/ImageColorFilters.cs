
using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;

namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class ImageColorFilters
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting(); 

            // Open the document.
            Document doc = new Document(string.Format("{0}{1}", dataDir, "TestFile.Colors.docx"));

            SaveColorTIFFwithLZW(doc, dataDir, 0.8f, 0.8f);
            SaveGrayscaleTIFFwithLZW(doc, dataDir, 0.8f, 0.8f);
            SaveBlackWhiteTIFFwithLZW(doc, dataDir, true);
            SaveBlackWhiteTIFFwithCITT4(doc, dataDir, true);
            SaveBlackWhiteTIFFwithRLE(doc, dataDir, true);
           
        }


        private static void SaveColorTIFFwithLZW(Document doc, string dataDir, float brightness, float contrast)
        {
            // Select the TIFF format with 100 dpi.
            ImageSaveOptions imgOpttiff = new ImageSaveOptions(SaveFormat.Tiff);
            imgOpttiff.Resolution = 100;

            // Select fullcolor LZW compression.
            imgOpttiff.TiffCompression = TiffCompression.Lzw;

            // Set brightness and contrast.
            imgOpttiff.ImageBrightness = brightness;
            imgOpttiff.ImageContrast = contrast;

            // Save multipage color TIFF.
            doc.Save(string.Format("{0}{1}", dataDir, "Result Colors.tiff"), imgOpttiff);

            Console.WriteLine("\nDocument converted to TIFF successfully with Colors.\nFile saved at " + dataDir + "Result Colors.tiff");
        }

        private static void SaveGrayscaleTIFFwithLZW(Document doc, string dataDir, float brightness, float contrast)
        {
            // Select the TIFF format with 100 dpi.
            ImageSaveOptions imgOpttiff = new ImageSaveOptions(SaveFormat.Tiff);
            imgOpttiff.Resolution = 100;

            // Select LZW compression.
            imgOpttiff.TiffCompression = TiffCompression.Lzw;

            // Apply grayscale filter.
            imgOpttiff.ImageColorMode = ImageColorMode.Grayscale;

            // Set brightness and contrast.
            imgOpttiff.ImageBrightness = brightness;
            imgOpttiff.ImageContrast = contrast;

            // Save multipage grayscale TIFF.
            doc.Save(string.Format("{0}{1}", dataDir, "Result Grayscale.tiff"), imgOpttiff);

            Console.WriteLine("\nDocument converted to TIFF successfully with Gray scale.\nFile saved at " + dataDir + "Result Grayscale.tiff");
        }

        private static void SaveBlackWhiteTIFFwithLZW(Document doc, string dataDir, bool highSensitivity)
        {
            // Select the TIFF format with 100 dpi.
            ImageSaveOptions imgOpttiff = new ImageSaveOptions(SaveFormat.Tiff);
            imgOpttiff.Resolution = 100;

            // Apply black & white filter. Set very high sensitivity to gray color.
            imgOpttiff.TiffCompression = TiffCompression.Lzw;
            imgOpttiff.ImageColorMode = ImageColorMode.BlackAndWhite;

            // Set brightness and contrast according to sensitivity.
            if (highSensitivity)
            {
                imgOpttiff.ImageBrightness = 0.4f;
                imgOpttiff.ImageContrast = 0.3f;
            }
            else
            {
                imgOpttiff.ImageBrightness = 0.9f;
                imgOpttiff.ImageContrast = 0.9f;
            }

            // Save multipage TIFF.
            doc.Save(string.Format("{0}{1}", dataDir, "result black and white.tiff"), imgOpttiff);

            Console.WriteLine("\nDocument converted to TIFF successfully with black and white.\nFile saved at " + dataDir + "Result black and white.tiff");
        }

        private static void SaveBlackWhiteTIFFwithCITT4(Document doc, string dataDir, bool highSensitivity)
        {
            // Select the TIFF format with 100 dpi.
            ImageSaveOptions imgOpttiff = new ImageSaveOptions(SaveFormat.Tiff);
            imgOpttiff.Resolution = 100;

            // Set CCITT4 compression.
            imgOpttiff.TiffCompression = TiffCompression.Ccitt4;

            // Apply grayscale filter.
            imgOpttiff.ImageColorMode = ImageColorMode.Grayscale;

            // Set brightness and contrast according to sensitivity.
            if (highSensitivity)
            {
                imgOpttiff.ImageBrightness = 0.4f;
                imgOpttiff.ImageContrast = 0.3f;
            }
            else
            {
                imgOpttiff.ImageBrightness = 0.9f;
                imgOpttiff.ImageContrast = 0.9f;
            }

            // Save multipage TIFF.
            doc.Save(string.Format("{0}{1}", dataDir, "result Ccitt4.tiff"), imgOpttiff);

            Console.WriteLine("\nDocument converted to TIFF successfully with black and white and Ccitt4 compression.\nFile saved at " + dataDir + "Result Ccitt4.tiff");
        }

        private static void SaveBlackWhiteTIFFwithRLE(Document doc, string dataDir, bool highSensitivity)
        {
            // Select the TIFF format with 100 dpi.
            ImageSaveOptions imgOpttiff = new ImageSaveOptions(SaveFormat.Tiff);
            imgOpttiff.Resolution = 100;

            // Set RLE compression.
            imgOpttiff.TiffCompression = TiffCompression.Rle;

            // Aply grayscale filter.
            imgOpttiff.ImageColorMode = ImageColorMode.Grayscale;

            // Set brightness and contrast according to sensitivity.
            if (highSensitivity)
            {
                imgOpttiff.ImageBrightness = 0.4f;
                imgOpttiff.ImageContrast = 0.3f;
            }
            else
            {
                imgOpttiff.ImageBrightness = 0.9f;
                imgOpttiff.ImageContrast = 0.9f;
            }

            // Save multipage TIFF grayscale with low bright and contrast
            doc.Save(string.Format("{0}{1}", dataDir, "result Rle.tiff"), imgOpttiff);

            Console.WriteLine("\nDocument converted to TIFF successfully with black and white and Rle compression.\nFile saved at " + dataDir + "Result Rle.tiff");
        }
    }
}