using Aspose.Words;
using Aspose.Words.Drawing;

namespace Image_ReSize
{
    class Program
    {
        static void Main(string[] args)
        {
			// define document file locaiton
			string fileDir = "../../data/Images.doc";

			Document doc = new Document();

			// load the document from disk
			DocumentBuilder builder = new DocumentBuilder(doc);

			builder.Write("Image Before ReSize");

			//insert image from disk
			Shape shape = builder.InsertImage(@"../../data/aspose_Words-for-net.jpg");

			// write text in document
            builder.Write("ReSize Image");

			//insert image from disk for resize
			shape = builder.InsertImage(@"../../data/aspose_Words-for-net.jpg");

			// To change the shape size. ( ConvertUtil Provides helper functions to convert between various measurement units. like Converts inches to points.)

			shape.Width = ConvertUtil.InchToPoint(0.5);
			shape.Height = ConvertUtil.InchToPoint(0.5);

			// save document on file location 
			builder.Document.Save("ImageReSize.doc");

        }
    }
}
