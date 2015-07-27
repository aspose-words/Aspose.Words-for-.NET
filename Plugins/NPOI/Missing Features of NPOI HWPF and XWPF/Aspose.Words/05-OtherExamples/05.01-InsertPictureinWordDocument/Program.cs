using System;
using System.Collections.Generic;
using System.Text; using Aspose.Words;

namespace _05._01_InsertPictureinWordDocument
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document("../../data/document.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.InsertImage("../../data/HumpbackWhale.jpg");

            doc.Save("insertedImage.docx");
        }
    }
}
