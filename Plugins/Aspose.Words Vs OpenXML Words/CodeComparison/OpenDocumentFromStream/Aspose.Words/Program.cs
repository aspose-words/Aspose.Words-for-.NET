using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words
{
    class Program
    {
        private static string strDoc = @"E:\Aspose\Aspose Vs OpenXML\Sample Files\OpenDocumentFromStream.docx";
        static void Main(string[] args)
        {
            string txt = "Append text in body - OpenAndAddToWordprocessingStream";
            Stream stream = File.Open(strDoc, FileMode.Open);
            OpenAndAddToWordprocessingStream(stream, txt);
        }

        private static void OpenAndAddToWordprocessingStream(Stream stream, string txt)
        {
            Document doc = new Document(stream);
            stream.Close();
            DocumentBuilder db = new DocumentBuilder(doc);
            db.Writeln(txt);
            doc.Save(strDoc);
        }
    }
}
