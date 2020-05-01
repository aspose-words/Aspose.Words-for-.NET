using System.IO;
using ApiExamples.TestData.TestClasses;
#if NET462 || JAVA
using System.Drawing;
#elif NETCOREAPP2_1 || __MOBILE__
using SkiaSharp;
#endif

namespace ApiExamples.TestData.TestBuilders
{
    public class ImageTestBuilder : ApiExampleBase
    {
#if NET462 || JAVA
        private Image mImage;
#elif NETCOREAPP2_1 || __MOBILE__
        private SKBitmap mImage;
#endif
        private Stream mImageStream;
        private byte[] mImageBytes;
        private string mImageString;

        public ImageTestBuilder()
        {
#if NET462 || JAVA
            mImage = Image.FromFile(ImageDir + "Transparent background logo.png");            
#elif NETCOREAPP2_1 || __MOBILE__
        this.mImage = SKBitmap.Decode(ImageDir + "Transparent background logo.png");
#endif
            mImageStream = Stream.Null;
            mImageBytes = new byte[0];
            mImageString = string.Empty;
        }

#if NET462 || JAVA
        public ImageTestBuilder WithImage(Image image)
        {
            mImage = image;
            return this;
        }
#elif NETCOREAPP2_1 || __MOBILE__
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