using System.IO;
#if NETSTANDARD2_0
using SkiaSharp;
#endif
#if !NETSTANDARD2_0
using System.Drawing;
#endif

namespace ApiExamples.TestData.TestClasses
{
    public class ImageTestClass
    {
#if NETSTANDARD2_0
        public SKBitmap Image { get; set; }
#else
        public Image Image { get; set; }
#endif
        public Stream ImageStream { get; set; }
        public byte[] ImageBytes { get; set; }
        public string ImageUri { get; set; }

#if NETSTANDARD2_0
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
            this.Image = image;
            this.ImageStream = imageStream;
            this.ImageBytes = imageBytes;
            this.ImageUri = imageUri;
        }
#endif
    }
}
