using System.IO;
using ApiExamples.TestData.TestClasses;
using System.Drawing;
#if NET5_0_OR_GREATER
using Image = SkiaSharp.SKBitmap;
#endif

namespace ApiExamples.TestData.TestBuilders
{
    public class ImageTestBuilder : ApiExampleBase
    {
        private Image mImage;
        private Stream mImageStream;
        private byte[] mImageBytes;
        private string mImageString;

        public ImageTestBuilder()
        {
#if NET48 || JAVA
            mImage = Image.FromFile(ImageDir + "Transparent background logo.png");
#elif NET5_0_OR_GREATER || __MOBILE__
            mImage = Image.Decode(ImageDir + "Transparent background logo.png");
#endif
            mImageStream = Stream.Null;
            mImageBytes = new byte[0];
            mImageString = string.Empty;
        }

        public ImageTestBuilder WithImage(string imagePath)
        {
#if NET48 || JAVA
            mImage = Image.FromFile(imagePath);
#elif NET5_0_OR_GREATER || __MOBILE__
            mImage = Image.Decode(imagePath);
#endif
            return this;
        }

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