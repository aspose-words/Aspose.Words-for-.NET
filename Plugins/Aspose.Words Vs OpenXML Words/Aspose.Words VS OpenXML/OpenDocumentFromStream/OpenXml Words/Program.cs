using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;

namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        private static string FilePath = @"..\..\..\..\Sample Files\";
        private static string fileName = FilePath + "OpenDocumentFromStream.docx";
        
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
