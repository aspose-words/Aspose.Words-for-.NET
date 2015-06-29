using Aspose.Words;
namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {
            string mypath = "";
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            // Create the headers.
            builder.MoveToHeaderFooter(HeaderFooterType.HeaderPrimary);
            builder.Write("Header Text goes here...");
            //add footer having current date
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);
            builder.InsertField("Date", "");

            doc.UpdateFields();
            doc.Save(mypath + "Insert Headers and Footers.doc");
        }
    }
}
