using System.Drawing;
using System.IO;

namespace ApiExamples.TestData
{
    public class ImageDataSource
    {
        public ImageDataSource(Stream stream)
        {
            this.Stream = stream;
        }

        public ImageDataSource(Image imageObject)
        {
            this.Image = imageObject;
        }

        public ImageDataSource(byte[] imageBytes)
        {
            this.Bytes = imageBytes;
        }

        public ImageDataSource(string uriToImage)
        {
            this.Uri = uriToImage;
        }

        public Stream Stream { get; set; }

        public Image Image { get; set; }

        public byte[] Bytes { get; set; }

        public string Uri { get; set; }
    }
}