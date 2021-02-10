// Copyright (c) 2001-2021 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using NUnit.Framework;
#if NET462 || JAVA
using System.Drawing.Printing;
using System.Windows.Forms;
using Aspose.Words;
using Aspose.Words.Rendering;
#endif

namespace ApiExamples
{
    [TestFixture]
    public class ExPrinting : ApiExampleBase
    {
#if NET462 || JAVA
        [Test, Ignore("Run only when the printer driver is installed.")]
        public void CustomPrint()
        {
            //ExStart
            //ExFor:PageInfo.GetDotNetPaperSize
            //ExFor:PageInfo.Landscape
            //ExSummary:Shows how to customize the printing of Aspose.Words documents.
            Document doc = new Document(MyDir + "Rendering.docx");

            MyPrintDocument printDoc = new MyPrintDocument(doc);
            printDoc.PrinterSettings.PrintRange = System.Drawing.Printing.PrintRange.SomePages;
            printDoc.PrinterSettings.FromPage = 1;
            printDoc.PrinterSettings.ToPage = 1;

            printDoc.Print();
        }

        /// <summary>
        /// Selects an appropriate paper size, orientation, and paper tray when printing.
        /// </summary>
        public class MyPrintDocument : PrintDocument
        {
            public MyPrintDocument(Document document)
            {
                mDocument = document;
            }

            /// <summary>
            /// Initializes the range of pages to be printed according to the user selection.
            /// </summary>
            protected override void OnBeginPrint(PrintEventArgs e)
            {
                base.OnBeginPrint(e);

                switch (PrinterSettings.PrintRange)
                {
                    case System.Drawing.Printing.PrintRange.AllPages:
                        mCurrentPage = 1;
                        mPageTo = mDocument.PageCount;
                        break;
                    case System.Drawing.Printing.PrintRange.SomePages:
                        mCurrentPage = PrinterSettings.FromPage;
                        mPageTo = PrinterSettings.ToPage;
                        break;
                    default:
                        throw new InvalidOperationException("Unsupported print range.");
                }
            }

            /// <summary>
            /// Called before each page is printed. 
            /// </summary>
            protected override void OnQueryPageSettings(QueryPageSettingsEventArgs e)
            {
                base.OnQueryPageSettings(e);

                // A single Microsoft Word document can have multiple sections that specify pages with different sizes, 
                // orientations, and paper trays. The .NET printing framework calls this code before 
                // each page is printed, which gives us a chance to specify how to print the current page.
                PageInfo pageInfo = mDocument.GetPageInfo(mCurrentPage - 1);
                e.PageSettings.PaperSize = pageInfo.GetDotNetPaperSize(PrinterSettings.PaperSizes);

                // Microsoft Word stores the paper source (printer tray) for each section as a printer-specific value.
                // To obtain the correct tray value, you will need to use the "RawKind" property, which your printer should return.
                e.PageSettings.PaperSource.RawKind = pageInfo.PaperTray;
                e.PageSettings.Landscape = pageInfo.Landscape;
            }

            /// <summary>
            /// Called for each page to render it for printing. 
            /// </summary>
            protected override void OnPrintPage(PrintPageEventArgs e)
            {
                base.OnPrintPage(e);

                // Aspose.Words rendering engine creates a page drawn from the origin (x = 0, y = 0) of the paper.
                // There will be a hard margin in the printer, which will render each page. We need to offset by that hard margin.
                float hardOffsetX, hardOffsetY;

                // Below are two ways of setting a hard margin.
                if (e.PageSettings != null && e.PageSettings.HardMarginX != 0 && e.PageSettings.HardMarginY != 0)
                {
                    // 1 -  Via the "PageSettings" property.
                    hardOffsetX = e.PageSettings.HardMarginX;
                    hardOffsetY = e.PageSettings.HardMarginY;
                }
                else
                {
                    // 2 -  Using our own values, if the "PageSettings" property is unavailable.
                    hardOffsetX = 20;
                    hardOffsetY = 20;
                }

                mDocument.RenderToScale(mCurrentPage, e.Graphics, -hardOffsetX, -hardOffsetY, 1.0f);

                mCurrentPage++;
                e.HasMorePages = mCurrentPage <= mPageTo;
            }

            private readonly Document mDocument;
            private int mCurrentPage;
            private int mPageTo;
        }
        //ExEnd

        [Test, Ignore("Run only when the printer driver is installed.")]
        public void PrintPageInfo()
        {
            //ExStart
            //ExFor:PageInfo
            //ExFor:PageInfo.GetSizeInPixels(Single, Single, Single)
            //ExFor:PageInfo.GetSpecifiedPrinterPaperSource(PaperSourceCollection, PaperSource)
            //ExFor:PageInfo.HeightInPoints
            //ExFor:PageInfo.Landscape
            //ExFor:PageInfo.PaperSize
            //ExFor:PageInfo.PaperTray
            //ExFor:PageInfo.SizeInPoints
            //ExFor:PageInfo.WidthInPoints
            //ExSummary:Shows how to print page size and orientation information for every page in a Word document.
            Document doc = new Document(MyDir + "Rendering.docx");

            // The first section has 2 pages. We will assign a different printer paper tray to each one,
            // whose number will match a kind of paper source. These sources and their Kinds will vary
            // depending on the installed printer driver.
            PrinterSettings.PaperSourceCollection paperSources = new PrinterSettings().PaperSources;

            doc.FirstSection.PageSetup.FirstPageTray = paperSources[0].RawKind;
            doc.FirstSection.PageSetup.OtherPagesTray = paperSources[1].RawKind;

            Console.WriteLine("Document \"{0}\" contains {1} pages.", doc.OriginalFileName, doc.PageCount);

            float scale = 1.0f;
            float dpi = 96;

            for (int i = 0; i < doc.PageCount; i++)
            {
                // Each page has a PageInfo object, whose index is the respective page's number.
                PageInfo pageInfo = doc.GetPageInfo(i);

                // Print the page's orientation and dimensions.
                Console.WriteLine($"Page {i + 1}:");
                Console.WriteLine($"\tOrientation:\t{(pageInfo.Landscape ? "Landscape" : "Portrait")}");
                Console.WriteLine($"\tPaper size:\t\t{pageInfo.PaperSize} ({pageInfo.WidthInPoints:F0}x{pageInfo.HeightInPoints:F0}pt)");
                Console.WriteLine($"\tSize in points:\t{pageInfo.SizeInPoints}");
                Console.WriteLine($"\tSize in pixels:\t{pageInfo.GetSizeInPixels(1.0f, 96)} at {scale * 100}% scale, {dpi} dpi");

                // Print the source tray information.
                Console.WriteLine($"\tTray:\t{pageInfo.PaperTray}");
                PaperSource source = pageInfo.GetSpecifiedPrinterPaperSource(paperSources, paperSources[0]);
                Console.WriteLine($"\tSuitable print source:\t{source.SourceName}, kind: {source.Kind}");
            }
            //ExEnd
        }

        [Test, Ignore("Run only when the printer driver is installed.")]
        public void PrinterSettingsContainer()
        {
            //ExStart
            //ExFor:PrinterSettingsContainer
            //ExFor:PrinterSettingsContainer.#ctor(PrinterSettings)
            //ExFor:PrinterSettingsContainer.DefaultPageSettingsPaperSource
            //ExFor:PrinterSettingsContainer.PaperSizes
            //ExFor:PrinterSettingsContainer.PaperSources
            //ExSummary:Shows how to access and list your printer's paper sources and sizes.
            // The "PrinterSettingsContainer" contains a "PrinterSettings" object,
            // which contains unique data for different printer drivers.
            PrinterSettingsContainer container = new PrinterSettingsContainer(new PrinterSettings());

            Console.WriteLine($"This printer contains {container.PaperSources.Count} printer paper sources:");
            foreach (PaperSource paperSource in container.PaperSources)
            {
                bool isDefault = container.DefaultPageSettingsPaperSource.SourceName == paperSource.SourceName;
                Console.WriteLine($"\t{paperSource.SourceName}, " +
                                  $"RawKind: {paperSource.RawKind} {(isDefault ? "(Default)" : "")}");
            }

            // The "PaperSizes" property contains the list of paper sizes to instruct the printer to use.
            // Both the PrinterSource and PrinterSize contain a "RawKind" property,
            // which equates to a paper type listed on the PaperSourceKind enum.
            // If there is a paper source with the same "RawKind" value as that of the printing page,
            // the printer will print the page using the provided paper source and size.
            // Otherwise, the printer will default to the source designated by the "DefaultPageSettingsPaperSource" property.
            Console.WriteLine($"{container.PaperSizes.Count} paper sizes:");
            foreach (System.Drawing.Printing.PaperSize paperSize in container.PaperSizes)
            {
                Console.WriteLine($"\t{paperSize}, RawKind: {paperSize.RawKind}");
            }
            //ExEnd
        }

        [Test, Ignore("Run only when the printer driver is installed.")]
        public void Print()
        {
            //ExStart
            //ExFor:Document.Print
            //ExFor:Document.Print(String)
            //ExSummary:Shows how to print a document using the default printer.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello world!");

            // Below are two ways of printing our document.
            // 1 -  Print using the default printer:
            doc.Print();

            // 2 -  Specify a printer that we wish to print the document with by name:
            string myPrinter = PrinterSettings.InstalledPrinters[4];

            Assert.AreEqual("HPDAAB96 (HP ENVY 5000 series)", myPrinter);

            doc.Print(myPrinter);
            //ExEnd
        }

        [Test, Ignore("Run only when the printer driver is installed.")]
        public void PrintRange()
        {
            //ExStart
            //ExFor:Document.Print(PrinterSettings)
            //ExFor:Document.Print(PrinterSettings, String)
            //ExSummary:Shows how to print a range of pages.
            Document doc = new Document(MyDir + "Rendering.docx");

            // Create a "PrinterSettings" object to modify how we print the document.
            PrinterSettings printerSettings = new PrinterSettings();

            // Set the "PrintRange" property to "PrintRange.SomePages" to
            // tell the printer that we intend to print only some document pages.
            printerSettings.PrintRange = System.Drawing.Printing.PrintRange.SomePages;

            // Set the "FromPage" property to "1", and the "ToPage" property to "3" to print pages 1 through to 3.
            // Page indexing is 1-based.
            printerSettings.FromPage = 1;
            printerSettings.ToPage = 3;

            // Below are two ways of printing our document.
            // 1 -  Print while applying our printing settings:
            doc.Print(printerSettings);

            // 2 -  Print while applying our printing settings, while also
            // giving the document a custom name that we may recognize in the printer queue:
            doc.Print(printerSettings, "My rendered document");
            //ExEnd
        }

        [Test, Ignore("Run only when the printer driver is installed.")]
        public void PreviewAndPrint()
        {
            //ExStart
            //ExFor:AsposeWordsPrintDocument.#ctor(Document)
            //ExFor:AsposeWordsPrintDocument.CachePrinterSettings
            //ExSummary:Shows how to select a page range and a printer to print the document with, and then bring up a print preview.
            Document doc = new Document(MyDir + "Rendering.docx");

            PrintPreviewDialog previewDlg = new PrintPreviewDialog();

            // Call the "Show" method to get the print preview form to show on top.
            previewDlg.Show();

            // Initialize the Print Dialog with the number of pages in the document.
            PrintDialog printDlg = new PrintDialog();
            printDlg.AllowSomePages = true;
            printDlg.PrinterSettings.MinimumPage = 1;
            printDlg.PrinterSettings.MaximumPage = doc.PageCount;
            printDlg.PrinterSettings.FromPage = 1;
            printDlg.PrinterSettings.ToPage = doc.PageCount;

            if (!printDlg.ShowDialog().Equals(DialogResult.OK))
                return;

            // Create the "Aspose.Words" implementation of the .NET print document,
            // and then pass the printer settings from the dialog.
            AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc);
            awPrintDoc.PrinterSettings = printDlg.PrinterSettings;

            // Use the "CachePrinterSettings" method to reduce time of the first call of the "Print" method.
            awPrintDoc.CachePrinterSettings();

            // Call the "Hide", and then the "InvalidatePreview" methods to get the print preview to show on top.
            previewDlg.Hide();
            previewDlg.PrintPreviewControl.InvalidatePreview();

            // Pass the "Aspose.Words" print document to the .NET Print Preview dialog.
            previewDlg.Document = awPrintDoc;

            previewDlg.ShowDialog();
            //ExEnd
        }
#endif
    }
}