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
            string MyDir = @"..\..\..\..\..\Sample Files\";
            Document doc = new Document(MyDir + "MyDocument.odt");

            doc.Save(MyDir+ "ConvertingFromOdt.docx", SaveFormat.Docx);
        }

        /// <summary>
        /// Loads an OTT document from the local file system and saves it to the DOCX format in a different file.  
        /// </summary>
        public static void ConvertingFromOtt()
        {
            string MyDir = @"..\..\..\..\..\Sample Files\";
            Document doc = new Document(MyDir + "MyDocument.ott");

            doc.Save(MyDir + "ConvertingFromOtt.docx", SaveFormat.Docx);
        }

        /// <summary>
        /// Loads a DOCX document from the local file system and saves it to the ODT format in a different file.  
        /// </summary>
        public static void ConvertingToOdt()
        {
            string MyDir = @"..\..\..\..\..\Sample Files\";
            Document doc = new Document(MyDir + "MyDocument.docx");

            doc.Save(MyDir + "ConvertingToOdt.odt", SaveFormat.Odt);
        }
    }
}
