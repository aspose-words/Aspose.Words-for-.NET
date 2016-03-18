using Aspose.Words;

namespace Convert_WordPage_Document_to_MultipageTIFF
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileDir = "../../data/";
            // open the document 
            Document doc = new Document(fileDir + "test.doc");
            // Save the document as multipage TIFF.
            doc.Save("TestFile Out.tiff"); 
         
        }
    }
}
