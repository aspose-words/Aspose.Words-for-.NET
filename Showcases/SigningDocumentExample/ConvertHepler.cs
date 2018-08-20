using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Aspose.Words;

namespace SigningDocumentExample
{
    public class ConvertHepler
    {
        /// <summary>
        /// Converting image file to bytes array
        /// </summary>
        /// <param name="pathToImage">Path to image</param>
        public static byte[] ConverImageToByteArray(string pathToImage)
        {
            Image imageIn = Image.FromFile(pathToImage);

            MemoryStream stream = new MemoryStream();
            imageIn.Save(stream, ImageFormat.Png);

            return stream.ToArray();
        }

        /// <summary>
        /// Converting bytes array to Aspose.Words.Document
        /// </summary>
        /// <param name="documentArray">Bytes array of document</param>
        public static Document ConvertByteArrayToDocument(byte[] documentArray)
        {
            MemoryStream stream = new MemoryStream(documentArray);
            Document document = new Document(stream);

            return document;
        }

        /// <summary>
        /// Converting Aspose.Words.Document to bytes array
        /// </summary>
        /// <param name="document">Aspose.Words.Document</param>
        public static byte[] ConvertDocumentToByteArray(Document document)
        {
            MemoryStream documentArray = new MemoryStream();
            document.Save(documentArray, SaveFormat.Docx);

            return documentArray.ToArray();
        }
    }
}
