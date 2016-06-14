
using System.IO;
using Aspose.Words;
using System;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class RemoveFooters
    {
        public static void Run()
        {
            //ExStart:RemoveFooters
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();

            Document doc = new Document(dataDir + "HeaderFooter.RemoveFooters.doc");

            foreach (Section section in doc)
            {
                // Up to three different footers are possible in a section (for first, even and odd pages).
                // We check and delete all of them.
                HeaderFooter footer;

                footer = section.HeadersFooters[HeaderFooterType.FooterFirst];
                if (footer != null)
                    footer.Remove();

                // Primary footer is the footer used for odd pages.
                footer = section.HeadersFooters[HeaderFooterType.FooterPrimary];
                if (footer != null)
                    footer.Remove();

                footer = section.HeadersFooters[HeaderFooterType.FooterEven];
                if (footer != null)
                    footer.Remove();
            }
            dataDir = dataDir + "HeaderFooter.RemoveFooters_out_.doc";

            // Save the document.
            doc.Save(dataDir);
            //ExEnd:RemoveFooters
            Console.WriteLine("\nAll footers from all sections deleted successfully.\nFile saved at " + dataDir);
        }
    }
}
