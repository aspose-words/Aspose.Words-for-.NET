using System.IO;
using ApiExamples.TestData.TestClasses;
#if NETFRAMEWORK
using System.Drawing;
#else
using SkiaSharp;
#endif

namespace ApiExamples.TestData.TestBuilders
{
    public class ImageTestBuilder : ApiExampleBase
    {
#if NETFRAMEWORK        
        private Image mImage;
#else
        private SKBitmap mImage;
#endif
        private Stream mImageStream;
        private byte[] mImageBytes;
        private string mImageString;

        public ImageTestBuilder()
        {
#if NETFRAMEWORK
            mImage = Image.FromFile(ImageDir + "Transparent background logo.png");            
#else
            this.mImage = SKBitmap.Decode(ImageDir + "Transparent background logo.png");
#endif
            mImageStream = Stream.Null;
            mImageBytes = new byte[0];
            mImageString = string.Empty;
        }

#if NETFRAMEWORK        
        public ImageTestBuilder WithImage(Image image)
        {
            mImage = image;
            return this;
        }
#else
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