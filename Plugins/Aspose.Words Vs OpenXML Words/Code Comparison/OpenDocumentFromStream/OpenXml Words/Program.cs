using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXml.Words
{
    class Program
    {
        private static string fileName = @"E:\Aspose\Aspose Vs OpenXML\Sample Files\OpenDocumentFromStream.docx";
        static void Main(string[] args)
        {
            string txt = "Append text in body - OpenAndAddToWordprocessingStream";
            Stream stream = File.Open(fileName, FileMode.Open);
            OpenAndAddToWordprocessingStream(stream, txt);
            stream.Close();
        }

        private static void OpenAndAddToWordprocessingStream(Stream stream, string txt)
        {
            WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(stream, true);
            Body body = wordprocessingDocument.MainDocumentPart.Document.Body;
            Paragraph para = body.AppendChild(new Paragraph());
            Run run = para.AppendChild(new Run());
            run.AppendChild(new Text(txt));
            wordprocessingDocument.Close();
        }
    }
}
