using Aspose.Words;
using System;

namespace Detect_the_File_Format
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"..\..\..\..\Sample Files\";

            FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath + "MyDocument.docx");

            Console.WriteLine("The document format is: " + FileFormatUtil.LoadFormatToExtension(info.LoadFormat));
            Console.WriteLine("Document is encrypted: " + info.IsEncrypted);
            Console.WriteLine("Document has a digital signature: " + info.HasDigitalSignature);
        }
    }
}
