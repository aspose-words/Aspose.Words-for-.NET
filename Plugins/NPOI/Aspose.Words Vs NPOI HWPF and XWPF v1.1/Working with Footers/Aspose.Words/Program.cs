using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
            Document wordDocument = new Document("data/Convert Word Doc to Other Formats.doc");
            HeaderFooterCollection footers = wordDocument.FirstSection.HeadersFooters;
            foreach (HeaderFooter footer in footers)
            {
                if (footer.HeaderFooterType == HeaderFooterType.FooterFirst || footer.HeaderFooterType == HeaderFooterType.FooterPrimary || footer.HeaderFooterType == HeaderFooterType.FooterEven)
                    Console.WriteLine(footer.GetText());
            }

        }
    }
}
