using Aspose.Words;

namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {
            string mypath = "";
            Document doc = new Document(mypath + "Search and Replace.doc");
            doc.Range.Replace("find me", "found", false, true);
            doc.Save(mypath + "Search and Replace.doc");
 
        }
    }
}
