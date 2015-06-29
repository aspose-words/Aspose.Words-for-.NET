using Aspose.Words;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {
            // The document that the other documents will be appended to.
            Document dstDoc = new Document();

            // We should call this method to clear this document of any existing content.
            dstDoc.RemoveAllChildren();

            int recordCount = 1;
            for (int i = 1; i <= recordCount; i++)
            {
                // Open the document to join.
                Document srcDoc = new Document("src.doc");

                // Append the source document at the end of the destination document.
                dstDoc.AppendDocument(srcDoc, ImportFormatMode.UseDestinationStyles);
                Document doc2 = new Document("Section.ModifyPageSetupInAllSections.doc");
                dstDoc.AppendDocument(doc2, ImportFormatMode.UseDestinationStyles);
                // In automation you were required to insert a new section break at this point, however in Aspose.Words we
                // don't need to do anything here as the appended document is imported as separate sectons already.

                // If this is the second document or above being appended then unlink all headers footers in this section
                // from the headers and footers of the previous section.
                if (i > 1)
                    dstDoc.Sections[i].HeadersFooters.LinkToPrevious(false);
            }
            dstDoc.Save("updated.doc");
        }
    }
}
