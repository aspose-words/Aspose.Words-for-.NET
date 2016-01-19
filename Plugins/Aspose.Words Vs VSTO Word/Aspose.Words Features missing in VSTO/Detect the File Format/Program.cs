using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Detect_the_File_Format
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = @"E:\Aspose\Aspose Vs VSTO\Aspose.Words Features missing in VSTO 1.1\Sample Files\";

            FileFormatInfo info = FileFormatUtil.DetectFileFormat(MyDir + "Detect_the_File_Format.doc");
            Console.WriteLine("The document format is: " + FileFormatUtil.LoadFormatToExtension(info.LoadFormat));
            Console.WriteLine("Document is encrypted: " + info.IsEncrypted);
            Console.WriteLine("Document has a digital signature: " + info.HasDigitalSignature);
        }
    }
}
