using Aspose.Words;

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

        /// <summary>
        /// Loads an ODT document from the local file system and saves it to the DOCX format in a different file.  
        /// </summary>
        public static void ConvertingFromOdt()
        {
            string MyDir = @"Files\";
            Document doc = new Document(MyDir + "OpenOfficeWord.odt");

            doc.Save(MyDir+"ConvertedOdtFromDoc.docx", SaveFormat.Docx);
        }

        /// <summary>
        /// Loads an OTT document from the local file system and saves it to the DOCX format in a different file.  
        /// </summary>
        public static void ConvertingFromOtt()
        {
            string MyDir = @"Files\";
            Document doc = new Document(MyDir + "Sample.ott");

            doc.Save(MyDir + "ConvertedFromOttFromDoc.docx", SaveFormat.Docx);
        }

        /// <summary>
        /// Loads a DOCX document from the local file system and saves it to the ODT format in a different file.  
        /// </summary>
        public static void ConvertingToOdt()
        {
            string MyDir = @"Files\";
            Document doc = new Document(MyDir + "ConvertedOdtFromDoc.docx");

            doc.Save(MyDir + "ConvertedToODT.odt", SaveFormat.Odt);
        }
    }
}
