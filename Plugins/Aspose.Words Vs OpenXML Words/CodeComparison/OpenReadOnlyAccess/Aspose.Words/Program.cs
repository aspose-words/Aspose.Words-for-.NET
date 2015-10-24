using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words
{
    class Program
    {
        private static string fileName = @"E:\Aspose\Aspose Vs OpenXML\Sample Files\OpenReadOnly.docx";
        static void Main(string[] args)
        {
            OpenWordprocessingDocumentReadonly(fileName);
        }

        private static void OpenWordprocessingDocumentReadonly(string fileName)
        {
            Document doc = new Document(fileName, new LoadOptions("1234"));
            DocumentBuilder db = new DocumentBuilder(doc);
            string txt = "Append text in body - OpenAndAddToWordprocessingStream";
            db.Writeln(txt);
            doc.Save(fileName);
        }
    }
}
