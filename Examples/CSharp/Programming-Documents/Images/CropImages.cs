using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Images
{
    class CropImages
    {
        public static void Run()
        {
            // ExStart:CropImageCall
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithImages();
            string inputPath = dataDir + "ch63_Fig0013.jpg";
            string outputPath = dataDir + "cropped-1.jpg";

            CropImage(inputPath,outputPath, 124, 90, 570, 571);
            // ExEnd:CropImageCall
            Console.WriteLine("\nCropped Image saved successfully.\nFile saved at " + outputPath);
        }
        // ExStart:CropImage
        public static void CropImage(string inPath, string outPath, int left, int top,int width, int height)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Image img = Image.FromFile(inPath);

            int effectiveWidth = img.Width - width;
            int effectiveHeight = img.Height - height;

            Shape croppedImage = builder.InsertImage(img,
                ConvertUtil.PixelToPoint(img.Width - effectiveWidth),
                ConvertUtil.PixelToPoint(img.Height - effectiveHeight));

            double widthRatio = croppedImage.Width / ConvertUtil.PixelToPoint(img.Width);
            double heightRatio = croppedImage.Height / ConvertUtil.PixelToPoint(img.Height);

            if (widthRatio< 1)
                croppedImage.ImageData.CropRight = 1 - widthRatio;

            if (heightRatio< 1)
                croppedImage.ImageData.CropBottom = 1 - heightRatio;

                float leftToWidth = (float)left / img.Width;
                float topToHeight = (float)top / img.Height;

                croppedImage.ImageData.CropLeft = leftToWidth;
                croppedImage.ImageData.CropRight = croppedImage.ImageData.CropRight - leftToWidth;

                croppedImage.ImageData.CropTop = topToHeight;
                croppedImage.ImageData.CropBottom = croppedImage.ImageData.CropBottom - topToHeight;

                croppedImage.GetShapeRenderer().Save(outPath, new ImageSaveOptions(SaveFormat.Jpeg));
        }
        // ExEnd:CropImage
    }
}
