using System;
using Aspose.Words;
using System.IO;

namespace _01._08_WorkingwithDigitalSignatures
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = "../../data/document.doc";

            FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath);
            
            if (info.HasDigitalSignature)
            {
                Console.WriteLine($"Document {new FileInfo(filePath).Name} has digital signatures, they will be lost if you open/save this document with Aspose.Words.");
            }
            else
            {
                Console.WriteLine("Document has no digital signature.");
            }
        }
    }
}
