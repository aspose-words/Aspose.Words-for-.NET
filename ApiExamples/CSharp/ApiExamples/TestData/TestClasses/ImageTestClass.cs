using System.IO;
#if NETFRAMEWORK || JAVA
using System.Drawing;
#else
using SkiaSharp;
#endif

namespace ApiExamples.TestData.TestClasses
{
    public class ImageTestClass
    {
#if NETFRAMEWORK || JAVA
        public Image Image { get; set; }        
#else
        public SKBitmap Image { get; set; }
#endif
        public Stream ImageStream { get; set; }
        public byte[] ImageBytes { get; set; }
        public string ImageUri { get; set; }

#if NETFRAMEWORK || JAVA
        public ImageTestClass(Image image, Stream imageStream, byte[] imageBytes, string imageUri)
        {
            Image = image;
            ImageStream = imageStream;
            ImageBytes = imageBytes;
            ImageUri = imageUri;
        }
#else
        public ImageTestClass(SKBitmap image, Stream imageStream, byte[] imageBytes, string imageUri)
        {
            this.Image = image;
            this.ImageStream = imageStream;
            this.ImageBytes = imageBytes;
            this.ImageUri = imageUri;
        }        
#endif
    }
}