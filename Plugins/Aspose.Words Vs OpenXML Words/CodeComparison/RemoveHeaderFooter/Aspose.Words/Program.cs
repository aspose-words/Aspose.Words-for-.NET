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
            string strDoc = @"C:\Users\Madiha\Dropbox\Word documents\RemoveHeaderFooter.docx";
            Document doc = new Document(strDoc);
            foreach (Section section in doc)
            {

                section.HeadersFooters.RemoveAt(0);
                HeaderFooter footer;
                // Primary footer is the footer used for odd pages.
                footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                if (footer != null)
                    footer.Remove();
            }

            doc.Save(strDoc);
        }
    }
}
