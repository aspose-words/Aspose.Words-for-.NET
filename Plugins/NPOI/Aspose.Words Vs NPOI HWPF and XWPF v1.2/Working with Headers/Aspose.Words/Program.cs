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
            HeaderFooterCollection headers = wordDocument.FirstSection.HeadersFooters;
            foreach (HeaderFooter header in headers)
            {
                if (header.HeaderFooterType == HeaderFooterType.HeaderFirst || header.HeaderFooterType == HeaderFooterType.HeaderPrimary || header.HeaderFooterType == HeaderFooterType.HeaderEven)
                    Console.WriteLine(header.GetText());
            }
        }
    }
}
