using System.IO;
using ApiExamples.TestData.TestClasses;
#if NET48 || JAVA
using System.Drawing;
#elif NET5_0_OR_GREATER || __MOBILE__
using SkiaSharp;
#endif

namespace ApiExamples.TestData.TestBuilders
{
    public class ImageTestBuilder : ApiExampleBase
    {
#if NET48 || JAVA
        private Image mImage;
#elif NET5_0_OR_GREATER || __MOBILE__
        private SKBitmap mImage;
#endif
        private Stream mImageStream;
        private byte[] mImageBytes;
        private string mImageString;

        public ImageTestBuilder()
        {
#if NET48 || JAVA
            mImage = Image.FromFile(ImageDir + "Transparent background logo.png");            
#elif NET5_0_OR_GREATER || __MOBILE__
        this.mImage = SKBitmap.Decode(ImageDir + "Transparent background logo.png");
#endif
            mImageStream = Stream.Null;
            mImageBytes = new byte[0];
            mImageString = string.Empty;
        }

#if NET48 || JAVA
        public ImageTestBuilder WithImage(Image image)
        {
            mImage = image;
            return this;
        }
#elif NET5_0_OR_GREATER || __MOBILE__
        public ImageTestBuilder WithImage(SKBitmap image)
        {
            this.mImage = image;
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

        public ImageTestBuilder WithImageString(string imageString)
        {
            mImageString = imageString;
            return this;
        }

        public ImageTestClass Build()
        {
            return new ImageTestClass(mImage, mImageStream, mImageBytes, mImageString);
        }
    }
}