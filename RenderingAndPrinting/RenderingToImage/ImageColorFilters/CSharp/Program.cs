//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System;
using System.IO;
using System.Reflection;

using Aspose.Words;
using Aspose.Words.Saving;

namespace ImageColorFiltersExample
{
    public class Program
    {
        public static void Main()
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            // Open the document.
            Document doc = new Document(string.Format("{0}{1}", dataDir, "TestFile.docx"));

            SaveColorTIFFwithLZW(doc, dataDir, 0.8f, 0.8f);
            SaveGrayscaleTIFFwithLZW(doc, dataDir, 0.8f, 0.8f);
            SaveBlackWhiteTIFFwithLZW(doc, dataDir, true);
            SaveBlackWhiteTIFFwithCITT4(doc, dataDir, true);
            SaveBlackWhiteTIFFwithRLE(doc, dataDir, true);
        }

        //ExStart
        //ExFor:ImageSaveOptions.ImageContrast
        //ExFor:ImageSaveOptions.ImageBrightness
        //ExId:ImageColorFilters_tiff_lzw_color
        //ExSummary: Applies LZW compression, saves to color TIFF image with specified brightness and contrast.
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
            doc.Save(string.Format("{0}{1}", dataDir, "result.tiff"), imgOpttiff);
        }
        //ExEnd

        //ExStart
        //ExFor:ImageColorMode
        //ExId:ImageColorFilters_tiff_lzw_grayscale
        //ExSummary: Applies LZW compression, saves to grayscale TIFF image with specified brightness and contrast.
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
            doc.Save(string.Format("{0}{1}", dataDir, "result.tiff"), imgOpttiff);
        }
        //ExEnd

        //ExStart
        //ExId:ImageColorFilters_tiff_lzw_blackandwhite_sens
        //ExSummary: Applies LZW compression, saves to black & white TIFF image with specified sensitivity to gray color.
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
            doc.Save(string.Format("{0}{1}", dataDir, "result.tiff"), imgOpttiff);
        }
        //ExEnd

        //ExStart
        //ExId:ImageColorFilters_tiff_ccitt4_graysvale_sens
        //ExSummary: Applies CCITT4 compression, saves to black & white TIFF image with specified sensitivity to gray color.
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
            doc.Save(string.Format("{0}{1}", dataDir, "result.tiff"), imgOpttiff);
        }
        //ExEnd

        //ExStart
        //ExId:ImageColorFilters_tiff_rle_graysvale_sens
        //ExSummary: Applies RLE compression with specified sensitivity to gray color, saves to black & white TIFF image.
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
            doc.Save(string.Format("{0}{1}", dataDir, "result.tiff"), imgOpttiff);
        }
        //ExEnd
    }
}