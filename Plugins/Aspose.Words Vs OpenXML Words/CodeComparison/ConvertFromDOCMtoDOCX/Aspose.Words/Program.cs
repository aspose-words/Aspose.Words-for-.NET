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
        private static string fileName = @"E:\Aspose\Aspose Vs OpenXML\Sample Files\DOCMtoDOCX.docm";
        static void Main(string[] args)
        {
            ConvertDOCMtoDOCX(fileName);
        }

        private static void ConvertDOCMtoDOCX(string fileName)
        {
            Document doc = new Document(fileName);
            var newFileName = Path.ChangeExtension(fileName, ".docx");
            doc.Save(newFileName, SaveFormat.Docx);
        }
    }
}
