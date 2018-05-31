using System.IO;
using ApiExamples.TestData.TestClasses;
#if NETSTANDARD2_0
using SkiaSharp;
#endif
#if !NETSTANDARD2_0 
using System.Drawing;
#endif

namespace ApiExamples.TestData.TestBuilders
{
    public class ImageTestBuilder : ApiExampleBase
    {
#if NETSTANDARD2_0
        private SKBitmap mImage;
#else
        private Image mImage;
#endif
        private Stream mImageStream;
        private byte[] mImageBytes;
        private string mImageUri;

        public ImageTestBuilder()
        {
#if NETSTANDARD2_0
            this.mImage = SKBitmap.Decode(ImageDir + "Watermark.png");
#else
            this.mImage = Image.FromFile(ImageDir + "Watermark.png");
#endif
            this.mImageStream = Stream.Null;
            this.mImageBytes = new byte[0];
            this.mImageUri = string.Empty;
        }

#if NETSTANDARD2_0
        public ImageTestBuilder WithImage(SKBitmap image)
        {
            this.mImage = image;
            return this;
        }
#else
        public ImageTestBuilder WithImage(Image image)
        {
            this.mImage = image;
            return this;
        }
#endif

        public ImageTestBuilder WithImageStream(Stream imageStream)
        {
            this.mImageStream = imageStream;
            return this;
        }

        public ImageTestBuilder WithImageBytes(byte[] imageBytes)
        {
            this.mImageBytes = imageBytes;
            return this;
        }

        public ImageTestBuilder WithImageUri(string imageUri)
        {
            this.mImageUri = imageUri;
            return this;
        }

        public ImageTestClass Build()
        {
            return new ImageTestClass(mImage, mImageStream, mImageBytes, mImageUri);
        }
    }
}
