using Aspose.Words;
using Aspose.Words.Drawing;
namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertImage("Download.jpg",
                RelativeHorizontalPosition.Margin,
                100,
                RelativeVerticalPosition.Margin,
                100,
                200,
                100,
                WrapType.Square);
            doc.Save("Picture Out.doc");

        }
    }
}
