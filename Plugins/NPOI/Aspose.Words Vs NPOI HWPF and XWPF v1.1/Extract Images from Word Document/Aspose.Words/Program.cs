using Aspose.Words.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
            Document wordDocument = new Document("data/Extract Images from Word Document.doc");
            NodeCollection pictures = wordDocument.GetChildNodes(NodeType.Shape, true);
            int imageindex = 0;
            foreach (Shape shape in pictures)
            {
                if (shape.HasImage)
                {
                    string imageFileName = "data/Aspose_" + (imageindex++).ToString() + "_" + shape.Name;
                    shape.ImageData.Save(imageFileName);
                }
            }

        }
    }
}
