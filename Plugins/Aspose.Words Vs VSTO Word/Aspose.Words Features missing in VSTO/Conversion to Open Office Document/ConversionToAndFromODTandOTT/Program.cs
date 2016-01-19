using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConversionToAndFromODTandOTT
{
    class Program
    {
        static void Main(string[] args)
        {
            ConvertingFromOdt();
            ConvertingFromOtt();
            ConvertingToOdt();
        }
        public static void ConvertingFromOdt()
        {
            string MyDir = @"Files\";
            Document doc = new Document(MyDir+"OpenOfficeWord.odt");
            doc.Save(MyDir+"ConvertedOdtFromDoc.docx",SaveFormat.Docx);
        }
        public static void ConvertingFromOtt()
        {
            string MyDir = @"Files\";
            Document doc = new Document(MyDir + "Sample.ott");
            doc.Save(MyDir + "ConvertedFromOttFromDoc.docx", SaveFormat.Docx);
        }
        public static void ConvertingToOdt()
        {
            string MyDir = @"Files\";
            Document doc = new Document(MyDir + "ConvertedOdtFromDoc.docx");
            doc.Save(MyDir + "ConvertedToODT.odt", SaveFormat.Odt);
        }
    }
}
