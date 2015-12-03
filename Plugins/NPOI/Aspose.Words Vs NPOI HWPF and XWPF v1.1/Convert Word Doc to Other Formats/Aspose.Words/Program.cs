using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;

namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
            Document wordDocument = new Document("data/Convert Word Doc to Other Formats.doc");
            wordDocument.Save("data/Convert Word Doc to Other Formatsblank.docx", SaveFormat.Docx);
            wordDocument.Save("data/Convert Word Doc to Other Formatsblank.bmp", SaveFormat.Bmp);
            wordDocument.Save("data/Convert Word Doc to Other Formatsblank.html", SaveFormat.Html);
            wordDocument.Save("data/Convert Word Doc to Other Formatsblank.pdf", SaveFormat.Pdf);
            wordDocument.Save("data/Convert Word Doc to Other Formatsblank.text", SaveFormat.Text);
        }
    }
}
