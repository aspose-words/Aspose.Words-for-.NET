#if NETFRAMEWORK
using Aspose.Words.Rendering;
using System;
using System.Collections.Generic;
using System.Drawing.Printing;

namespace ApiExamples
{
    //ExStart:PrintTracker
    //GistId:571cc6e23284a2ec075d15d4c32e3bbf
    //ExFor:AsposeWordsPrintDocument
    //ExFor:AsposeWordsPrintDocument.PagesRemaining
    //ExSummary:Shows an example class for monitoring the progress of printing.
    /// <summary>
    /// Tracks printing progress of an Aspose.Words document and logs printing events.
    /// </summary>
    internal class PrintTracker
    {
        /// <summary>
        /// Initializes a new instance of the SamplePrintTracker class
        /// and subscribes to printing events of the specified document.
        /// </summary>
        /// <param name="printDoc">The Aspose.Words print document to track.</param>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="printDoc"/> is null.</exception>
        internal PrintTracker(AsposeWordsPrintDocument printDoc)
        {
            if (printDoc == null)
                throw new ArgumentNullException(nameof(printDoc));

            printDoc.BeginPrint += PrintDocument_BeginPrint;
            printDoc.EndPrint += PrintDocument_EndPrint;
            printDoc.PrintPage += PrintDocument_PrintPage;
        }

        /// <summary>
        /// Gets the current page being printed (1-based index).
        /// Returns -1 when no printing is in progress.
        /// </summary>
        internal int PrintingPage { get; private set; } = -1;

        /// <summary>
        /// Gets the total number of pages to print.
        /// Returns 0 when no printing is in progress.
        /// </summary>
        internal int TotalPages { get; private set; }

        /// <summary>
        /// Gets the log of printing events in chronological order.
        /// </summary>
        internal IReadOnlyList<string> EventLog { get; } = new List<string>();

        private void PrintDocument_BeginPrint(object sender, PrintEventArgs e)
        {
            var printDoc = (AsposeWordsPrintDocument)sender;

            PrintingPage = -1;
            TotalPages = printDoc.PagesRemaining;

            AddLogEntry($"BeginPrint. {printDoc.PagesRemaining} pages left to print.");
        }

        private void PrintDocument_PrintPage(object sender, PrintPageEventArgs e)
        {
            var printDoc = (AsposeWordsPrintDocument)sender;

            PrintingPage = TotalPages - printDoc.PagesRemaining + 1;

            AddLogEntry($"Printing page {PrintingPage} of {TotalPages}");
        }

        private void PrintDocument_EndPrint(object sender, PrintEventArgs e)
        {
            var printDoc = (AsposeWordsPrintDocument)sender;

            PrintingPage = -1;
            TotalPages = 0;

            AddLogEntry($"EndPrint. {printDoc.PagesRemaining} pages left to print.");
        }

        private void AddLogEntry(string message)
        {
            ((List<string>)EventLog).Add(message);
        }
    }
}
//ExEnd:PrintTracker
#endif