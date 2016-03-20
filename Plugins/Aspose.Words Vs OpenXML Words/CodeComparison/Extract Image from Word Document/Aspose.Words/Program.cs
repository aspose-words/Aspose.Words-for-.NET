using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using Aspose.Words.Drawing;
using System.IO;
namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = "test.docx";
            Document doc = new Document(fileName);
 
            // Save document as DOC in memory
            MemoryStream stream = new MemoryStream();
            doc.Save(stream, SaveFormat.Doc);
 
            // Reload document as DOC to extract images.
            Document doc2 = new Document(stream);
            NodeCollection shapes = doc2.GetChildNodes(NodeType.Shape, true);
            int imageIndex = 0;
            foreach (Shape shape in shapes)
            {
                if (shape.HasImage)
                {
                    string imageFileName = string.Format(
                        "Image.ExportImages.{0}_out_{1}", imageIndex, FileFormatUtil.ImageTypeToExtension(shape.ImageData.ImageType));
                    shape.ImageData.Save(imageFileName);
                    imageIndex++;
                }
            }
        }
    }
}
