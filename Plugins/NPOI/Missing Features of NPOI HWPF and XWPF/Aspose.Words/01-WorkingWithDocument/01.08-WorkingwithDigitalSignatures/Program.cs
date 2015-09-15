using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Words;
using System.IO;

namespace _01._08_WorkingwithDigitalSignatures
{
    class Program
    {
        static void Main(string[] args)
        {
            // The path to the document which is to be processed.

            string filePath = "../../data/document.doc";

            FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath);
            
            if (info.HasDigitalSignature)
            {
                Console.WriteLine(string.Format("Document {0} has digital signatures, they will be lost if you open/save this document with Aspose.Words.", new FileInfo(filePath).Name));
            }
            else
            {
                Console.WriteLine("Document has no digital signature.");
            }
        }
    }
}
