using System.IO;
using ApiExamples.TestData.TestClasses;
#if NETSTANDARD2_0 || __MOBILE__
using SkiaSharp;
#endif
#if !(NETSTANDARD2_0 || __MOBILE__)
using System.Drawing;

#endif

namespace ApiExamples.TestData.TestBuilders
{
    public class ImageTestBuilder : ApiExampleBase
    {
#if NETSTANDARD2_0 || __MOBILE__
        private SKBitmap mImage;
#else
        private Image mImage;
#endif
        private Stream mImageStream;
        private byte[] mImageBytes;
        private string mImageUri;

        public ImageTestBuilder()
        {
#if NETSTANDARD2_0 || __MOBILE__
            this.mImage = SKBitmap.Decode(ImageDir + "Watermark.png");
#else
            mImage = Image.FromFile(ImageDir + "Watermark.png");
#endif
            mImageStream = Stream.Null;
            mImageBytes = new byte[0];
            mImageUri = string.Empty;
        }

#if NETSTANDARD2_0 || __MOBILE__
        public ImageTestBuilder WithImage(SKBitmap image)
        {
            this.mImage = image;
            return this;
        }
#else
        public ImageTestBuilder WithImage(Image image)
        {
            mImage = image;
            return this;
        }
#endif

        public ImageTestBuilder WithImageStream(Stream imageStream)
        {
            mImageStream = imageStream;
            return this;
        }

        public ImageTestBuilder WithImageBytes(byte[] imageBytes)
        {
            mImageBytes = imageBytes;
            return this;
        }

        public ImageTestBuilder WithImageUri(string imageUri)
        {
            mImageUri = imageUri;
            return this;
        }

        public ImageTestClass Build()
        {
            return new ImageTestClass(mImage, mImageStream, mImageBytes, mImageUri);
        }
    }
}