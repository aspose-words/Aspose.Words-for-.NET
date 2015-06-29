using Aspose.Words;
namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {

            string MyDir = "";
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            //Add picture
            builder.InsertImage(MyDir + "download.jpg");
            
            doc.Save(MyDir+"Adding Picture.doc");

           
        }
    }
}
