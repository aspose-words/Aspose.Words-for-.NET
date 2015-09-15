using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string strDoc = @"C:\Users\Madiha\Dropbox\Word documents\RemoveHeaderFooter.docx";
            RemoveHeadersAndFooters(strDoc);

        }
        // Remove all of the headers and footers from a document.
        public static void RemoveHeadersAndFooters(string filename)
        {
            // Given a document name, remove all of the headers and footers
            // from the document.
            using (WordprocessingDocument doc =
                   WordprocessingDocument.Open(filename, true))
            {
                // Get a reference to the main document part.
                var docPart = doc.MainDocumentPart;
                // Count the header and footer parts and continue if there 
                // are any.
                if (docPart.HeaderParts.Count() > 0 ||
                    docPart.FooterParts.Count() > 0)
                {
                    // Remove the header and footer parts.
                    docPart.DeleteParts(docPart.HeaderParts);
                    docPart.DeleteParts(docPart.FooterParts);

                    // Get a reference to the root element of the main
                    // document part.
                    Document document = docPart.Document;

                    // Remove all references to the headers and footers.

                    // First, create a list of all descendants of type
                    // HeaderReference. Then, navigate the list and call
                    // Remove on each item to delete the reference.
                    var headers =
                      document.Descendants<HeaderReference>().ToList();
                    foreach (var header in headers)
                    {
                        header.Remove();
                    }

                    // First, create a list of all descendants of type
                    // FooterReference. Then, navigate the list and call
                    // Remove on each item to delete the reference.
                    var footers =
                      document.Descendants<FooterReference>().ToList();
                    foreach (var footer in footers)
                    {
                        footer.Remove();
                    }

                    // Save the changes.
                    document.Save();
                }
            }
        }
    }
}