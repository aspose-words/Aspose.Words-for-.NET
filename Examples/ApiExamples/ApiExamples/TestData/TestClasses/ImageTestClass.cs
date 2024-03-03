using System.IO;
using System.Drawing;
#if NET5_0_OR_GREATER || __MOBILE__
using Image = SkiaSharp.SKBitmap;
#endif

namespace ApiExamples.TestData.TestClasses
{
    public class ImageTestClass
    {
        public Image Image { get; set; }
        public Stream ImageStream { get; set; }
        public byte[] ImageBytes { get; set; }
        public string ImageString { get; set; }

        public ImageTestClass(Image image, Stream imageStream, byte[] imageBytes, string imageString)
        {
            Image = image;
            ImageStream = imageStream;
            ImageBytes = imageBytes;
            ImageString = imageString;
        }
    }
}