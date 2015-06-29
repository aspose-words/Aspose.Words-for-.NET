using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Drawing;
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
            doc.Save(MyDir+"Add Picture and WordArt.doc");

           
        }
    }
}
