using System.IO;
#if NETSTANDARD2_0 || __MOBILE__
using SkiaSharp;
#endif
#if !(NETSTANDARD2_0 || __MOBILE__)
using System.Drawing;

#endif

namespace ApiExamples.TestData.TestClasses
{
    public class ImageTestClass
    {
#if NETSTANDARD2_0 || __MOBILE__
        public SKBitmap Image { get; set; }
#else
        public Image Image { get; set; }
#endif
        public Stream ImageStream { get; set; }
        public byte[] ImageBytes { get; set; }
        public string ImageUri { get; set; }

#if NETSTANDARD2_0 || __MOBILE__
        public ImageTestClass(SKBitmap image, Stream imageStream, byte[] imageBytes, string imageUri)
        {
            this.Image = image;
            this.ImageStream = imageStream;
            this.ImageBytes = imageBytes;
            this.ImageUri = imageUri;
        }
#else
        public ImageTestClass(Image image, Stream imageStream, byte[] imageBytes, string imageUri)
        {
            Image = image;
            ImageStream = imageStream;
            ImageBytes = imageBytes;
            ImageUri = imageUri;
        }
#endif
    }
}