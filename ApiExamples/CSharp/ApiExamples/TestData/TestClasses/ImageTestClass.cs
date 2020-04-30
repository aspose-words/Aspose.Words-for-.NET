using System.IO;
#if NET462 || JAVA
using System.Drawing;
#elif NETCOREAPP2_1 || __MOBILE__
using SkiaSharp;
#endif

namespace ApiExamples.TestData.TestClasses
{
    public class ImageTestClass
    {
#if NET462 || JAVA
        public Image Image { get; set; }        
#elif NETCOREAPP2_1 || __MOBILE__
        public SKBitmap Image { get; set; }
#endif
        public Stream ImageStream { get; set; }
        public byte[] ImageBytes { get; set; }
        public string ImageString { get; set; }

#if NET462 || JAVA
        public ImageTestClass(Image image, Stream imageStream, byte[] imageBytes, string imageString)
        {
            Image = image;
            ImageStream = imageStream;
            ImageBytes = imageBytes;
            ImageString = imageString;
        }
#elif NETCOREAPP2_1 || __MOBILE__
        public ImageTestClass(SKBitmap image, Stream imageStream, byte[] imageBytes, string imageString)
        {
            this.Image = image;
            this.ImageStream = imageStream;
            this.ImageBytes = imageBytes;
            this.ImageString = imageString;
        }        
#endif
    }
}