using Aspose.Words;
namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {
            string mypath = "";
            Document doc = new Document(mypath + "Remove Headers and Footers.doc");
            foreach (Section section in doc)
            {
                
                section.HeadersFooters.RemoveAt(0);
                HeaderFooter footer;
                // Primary footer is the footer used for odd pages.
                footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                if (footer != null)
                    footer.Remove();
            }
            
            doc.Save(mypath + "Remove Headers and Footers.doc");
        }
    }
}
