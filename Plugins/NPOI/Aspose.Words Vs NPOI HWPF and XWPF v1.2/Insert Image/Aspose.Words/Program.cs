using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertImage("Image.png");
            doc.Save("Aspose_InsertImage.docx");

        }
    }
}
