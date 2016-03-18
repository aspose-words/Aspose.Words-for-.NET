using Aspose.Words;
using Aspose.Words.Saving;

namespace Convert_Doc_to_Png
{
    class Program
    {
        static void Main(string[] args)
        {
			// define document file location
            string fileDir = "../../data/";

            // load the document.
            Document doc = new Document(fileDir + "test.doc");

            //Create an ImageSaveOptions object to pass to the Save method
            ImageSaveOptions options = new ImageSaveOptions(SaveFormat.Png);

            // Save each page of the document as Png in data folder
            for (int i = 0; i < doc.PageCount; i++)
            {
                options.PageIndex = i;
                doc.Save(string.Format(i + "WordDocumentAsPNG out.Png", i), options);
            }
        }
    }
}
