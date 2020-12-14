using Aspose.Words.Replacing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Find_Replace
{
    class ReplaceInHeaderAndFooter
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_FindAndReplace();
        }

        private static void ReplaceTextInFooter(string dataDir)
        {
            // ExStart:ReplaceTextInFooter
            // Open the template document, containing obsolete copyright information in the footer.
            Document doc = new Document(dataDir + "HeaderFooter.ReplaceText.doc");

            // Access header of the Word document.
            HeaderFooterCollection headersFooters = doc.FirstSection.HeadersFooters;
            HeaderFooter header = headersFooters[HeaderFooterType.HeaderPrimary];

            // Set options.
            FindReplaceOptions options = new FindReplaceOptions
            {
                MatchCase = false,
                FindWholeWordsOnly = false
            };

            // Replace text in the header of the Word document.
            header.Range.Replace("Aspose.Words", "Remove", options);

            // Save the Word document.
            doc.Save(dataDir + "HeaderReplace.docx");
            // ExEnd:ReplaceTextInFooter
        }

        // ExStart:ShowChangesForHeaderAndFooterOrders
        private static void ShowChangesForHeaderAndFooterOrders(string dataDir)
        {
            Document doc = new Document(dataDir + "HeaderFooter.HeaderFooterOrder.docx");

            // Assert that we use special header and footer for the first page
            // The order for this: first header\footer, even header\footer, primary header\footer
            Section firstPageSection = doc.FirstSection;

            ReplaceLog logger = new ReplaceLog();
            FindReplaceOptions options = new FindReplaceOptions { ReplacingCallback = logger };

            doc.Range.Replace(new Regex("(header|footer)"), "", options);

            doc.Save(dataDir + "HeaderFooter.HeaderFooterOrder.docx");

            // Prepare our string builder for assert results without "DifferentFirstPageHeaderFooter"
            logger.ClearText();

            // Remove special first page
            // The order for this: primary header, default header, primary footer, default footer, even header\footer
            firstPageSection.PageSetup.DifferentFirstPageHeaderFooter = false;

            doc.Range.Replace(new Regex("(header|footer)"), "", options);
        }
        
        private class ReplaceLog : IReplacingCallback
        {
            public ReplaceAction Replacing(ReplacingArgs args)
            {
                _textBuilder.AppendLine(args.MatchNode.GetText());
                return ReplaceAction.Skip;
            }

            internal void ClearText()
            {
                _textBuilder.Clear();
            }

            internal string Text
            {
                get { return _textBuilder.ToString(); }
            }

            private readonly StringBuilder _textBuilder = new StringBuilder();
        }
        // ExEnd:ShowChangesForHeaderAndFooterOrders
    }
}
