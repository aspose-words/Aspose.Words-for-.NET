#if NET462
using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Aspose.Words;
using Aspose.Words.Rendering;
using NUnit.Framework;

namespace DocsExamples.Rendering_and_Printing
{
    internal class PrintDocuments : DocsExamplesBase
    {
        [Test, Ignore("Run only when the printer driver is installed")]
        public void CachePrinterSettings()
        {
            //ExStart:CachePrinterSettings
            Document doc = new Document(MyDir + "Rendering.docx");

            doc.UpdatePageLayout();

            PrinterSettings settings = new PrinterSettings { PrinterName = "Microsoft XPS Document Writer" };

            // The standard print controller comes with no UI.
            PrintController standardPrintController = new StandardPrintController();

            AsposeWordsPrintDocument printDocument = new AsposeWordsPrintDocument(doc)
            {
                PrinterSettings = settings,
                PrintController = standardPrintController
            };
            printDocument.CachePrinterSettings();

            printDocument.Print();
            //ExEnd:CachePrinterSettings
        }

        [Test, Ignore("Run only when the printer driver is installed")]
        public void Print()
        {
            Document doc = new Document(MyDir + "Rendering.docx");

            //ExStart:PrintDialog
            // Initialize the print dialog with the number of pages in the document.
            PrintDialog printDlg = new PrintDialog
            {
                AllowSomePages = true,
                PrinterSettings =
                {
                    MinimumPage = 1, MaximumPage = doc.PageCount, FromPage = 1, ToPage = doc.PageCount
                }
            };
            //ExEnd:PrintDialog

            //ExStart:ShowDialog
            if (printDlg.ShowDialog() != DialogResult.OK)
                return;
            //ExEnd:ShowDialog

            //ExStart:AsposeWordsPrintDocument
            // Pass the printer settings from the dialog to the print document.
            AsposeWordsPrintDocument awPrintDoc = new AsposeWordsPrintDocument(doc)
            {
                PrinterSettings = printDlg.PrinterSettings
            };
            //ExEnd:AsposeWordsPrintDocument

            //ExStart:ActivePrintPreviewDialog
            // Pass the Aspose.Words print document to the Print Preview dialog.
            ActivePrintPreviewDialog previewDlg = new ActivePrintPreviewDialog
            {
                Document = awPrintDoc, ShowInTaskbar = true, MinimizeBox = true
            };
            
            // Specify additional parameters of the Print Preview dialog.
            previewDlg.PrintPreviewControl.Zoom = 1;
            previewDlg.Document.DocumentName = "PrintDocuments.Print.docx";
            previewDlg.WindowState = FormWindowState.Maximized;
            previewDlg.ShowDialog(); // Show the appropriately configured Print Preview dialog.
            //ExEnd:ActivePrintPreviewDialog
        }

        [Test, Ignore("Run only when the printer driver is installed")]
        public void PrintMultiplePages()
        {
            //ExStart:PrintMultiplePagesOnOneSheet
            Document doc = new Document(MyDir + "Rendering.docx");

            //ExStart:PrintDialogSettings
            // Initialize the Print Dialog with the number of pages in the document.
            PrintDialog printDlg = new PrintDialog
            {
                AllowSomePages = true,
                PrinterSettings =
                {
                    MinimumPage = 1, MaximumPage = doc.PageCount, FromPage = 1, ToPage = doc.PageCount
                }
            };
            //ExEnd:PrintDialogSettings

            // Check if the user accepted the print settings and proceed to preview.
            //ExStart:CheckPrintSettings
            if (printDlg.ShowDialog() != DialogResult.OK)
                return;
            //ExEnd:CheckPrintSettings

            // Pass the printer settings from the dialog to the print document.
            MultipagePrintDocument awPrintDoc = new MultipagePrintDocument(doc, 4, true)
            {
                PrinterSettings = printDlg.PrinterSettings
            };

            //ExStart:ActivePrintPreviewDialog
            // Create and configure the the ActivePrintPreviewDialog class.
            ActivePrintPreviewDialog previewDlg = new ActivePrintPreviewDialog
            {
                Document = awPrintDoc, ShowInTaskbar = true, MinimizeBox = true
            };

            // Specify additional parameters of the Print Preview dialog.
            previewDlg.Document.DocumentName = "PrintDocuments.PrintMultiplePages.docx";
            previewDlg.WindowState = FormWindowState.Maximized;
            previewDlg.ShowDialog(); // Show appropriately configured Print Preview dialog.
            //ExEnd:ActivePrintPreviewDialog
            //ExEnd:PrintMultiplePagesOnOneSheet
        }

        [Test, Ignore("Run only when a printer driver installed")]
        public void UseXpsPrintHelper()
        {
            //ExStart:PrintDocViaXpsPrint
            Document document = new Document(MyDir + "Rendering.docx");

            // Specify the name of the printer you want to print to.
            const string printerName = @"\\COMPANY\Brother MFC-885CW Printer";

            XpsPrintHelper.Print(document, printerName, "My Test Job", true);
            //ExEnd:PrintDocViaXpsPrint
        }
    }

    //ExStart:MultipagePrintDocument
    internal class MultipagePrintDocument : PrintDocument
    //ExEnd:MultipagePrintDocument
    {
        //ExStart:DataAndStaticFields
        private readonly Document mDocument;
        private readonly int mPagesPerSheet;
        private readonly bool mPrintPageBorders;
        private Size mPaperSize;
        private int mCurrentPage;

        private int mPageTo;
        //ExEnd:DataAndStaticFields

        /// <summary>
        /// The constructor of the custom PrintDocument class.
        /// </summary> 
        //ExStart:MultipagePrintDocumentConstructor 
        public MultipagePrintDocument(Document document, int pagesPerSheet, bool printPageBorders)
        {
            mDocument = document ?? throw new ArgumentNullException(nameof(document));
            mPagesPerSheet = pagesPerSheet;
            mPrintPageBorders = printPageBorders;
        }
        //ExEnd:MultipagePrintDocumentConstructor

        /// <summary>
        /// The overridden method OnBeginPrint, which is called before the first page of the document prints.
        /// </summary>
        //ExStart:OnBeginPrint
        protected override void OnBeginPrint(PrintEventArgs e)
        {
            base.OnBeginPrint(e);

            switch (PrinterSettings.PrintRange)
            {
                case PrintRange.AllPages:
                    mCurrentPage = 0;
                    mPageTo = mDocument.PageCount - 1;
                    break;
                case PrintRange.SomePages:
                    mCurrentPage = PrinterSettings.FromPage - 1;
                    mPageTo = PrinterSettings.ToPage - 1;
                    break;
                default:
                    throw new InvalidOperationException("Unsupported print range.");
            }

            // Store the size of the paper selected by the user, taking into account the paper orientation.
            if (PrinterSettings.DefaultPageSettings.Landscape)
                mPaperSize = new Size(PrinterSettings.DefaultPageSettings.PaperSize.Height,
                    PrinterSettings.DefaultPageSettings.PaperSize.Width);
            else
                mPaperSize = new Size(PrinterSettings.DefaultPageSettings.PaperSize.Width,
                    PrinterSettings.DefaultPageSettings.PaperSize.Height);
        }
        //ExEnd:OnBeginPrint

        /// <summary>
        /// Generates the printed page from the specified number of the document pages.
        /// </summary>
        //ExStart:OnPrintPage
        protected override void OnPrintPage(PrintPageEventArgs e)
        {
            base.OnPrintPage(e);

            // Transfer to the point units.
            e.Graphics.PageUnit = GraphicsUnit.Point;
            e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAliasGridFit;

            // Get the number of the thumbnail placeholders across and down the paper sheet.
            Size thumbCount = GetThumbCount(mPagesPerSheet);

            // Calculate the size of each thumbnail placeholder in points.
            // Papersize in .NET is represented in hundreds of an inch. We need to convert this value to points first.
            SizeF thumbSize = new SizeF(
                HundredthsInchToPoint(mPaperSize.Width) / thumbCount.Width,
                HundredthsInchToPoint(mPaperSize.Height) / thumbCount.Height);

            // Select the number of the last page to be printed on this sheet of paper.
            int pageTo = Math.Min(mCurrentPage + mPagesPerSheet - 1, mPageTo);

            // Loop through the selected pages from the stored current page to the calculated last page.
            for (int pageIndex = mCurrentPage; pageIndex <= pageTo; pageIndex++)
            {
                // Calculate the column and row indices.
                int rowIdx = Math.DivRem(pageIndex - mCurrentPage, thumbCount.Width, out int columnIdx);

                // Define the thumbnail location in world coordinates (points in this case).
                float thumbLeft = columnIdx * thumbSize.Width;
                float thumbTop = rowIdx * thumbSize.Height;
                // Render the document page to the Graphics object using calculated coordinates and thumbnail placeholder size.
                // The useful return value is the scale at which the page was rendered.
                float scale = mDocument.RenderToSize(pageIndex, e.Graphics, thumbLeft, thumbTop, thumbSize.Width,
                    thumbSize.Height);

                // Draw the page borders (the page thumbnail could be smaller than the thumbnail placeholder size).
                if (mPrintPageBorders)
                {
                    // Get the real 100% size of the page in points.
                    SizeF pageSize = mDocument.GetPageInfo(pageIndex).SizeInPoints;
                    // Draw the border around the scaled page using the known scale factor.
                    e.Graphics.DrawRectangle(Pens.Black, thumbLeft, thumbTop, pageSize.Width * scale,
                        pageSize.Height * scale);

                    // Draws the border around the thumbnail placeholder.
                    e.Graphics.DrawRectangle(Pens.Red, thumbLeft, thumbTop, thumbSize.Width, thumbSize.Height);
                }
            }

            // Re-calculate the next current page and continue with printing if such a page resides within the print range.
            mCurrentPage += mPagesPerSheet;
            e.HasMorePages = mCurrentPage <= mPageTo;
        }
        //ExEnd:OnPrintPage

        /// <summary>
        /// Converts hundredths of inches to points.
        /// </summary>
        //ExStart:HundredthsInchToPoint
        private float HundredthsInchToPoint(float value)
        {
            return (float)ConvertUtil.InchToPoint(value / 100);
        }
        //ExEnd:HundredthsInchToPoint

        /// <summary>
        /// Defines the number of columns and rows depending on the pagesPerSheet number and the page orientation.
        /// </summary>
        //ExStart:GetThumbCount
        private Size GetThumbCount(int pagesPerSheet)
        {
            Size size;
            // Define the number of columns and rows on the sheet for the Landscape-oriented paper.
            switch (pagesPerSheet)
            {
                case 16:
                    size = new Size(4, 4);
                    break;
                case 9:
                    size = new Size(3, 3);
                    break;
                case 8:
                    size = new Size(4, 2);
                    break;
                case 6:
                    size = new Size(3, 2);
                    break;
                case 4:
                    size = new Size(2, 2);
                    break;
                case 2:
                    size = new Size(2, 1);
                    break;
                default:
                    size = new Size(1, 1);
                    break;
            }

            // Switch the width and height of the paper is in the Portrait orientation.
            return mPaperSize.Width < mPaperSize.Height ? new Size(size.Height, size.Width) : size;
        }
        //ExEnd:GetThumbCount
    }

    /// <summary>
    /// A utility class that converts a document to XPS using Aspose.Words and then sends to the XpsPrint API.
    /// </summary>
    public class XpsPrintHelper
    {
        /// <summary>
        /// No ctor.
        /// </summary>
        private XpsPrintHelper()
        {
        }

        //ExStart:XpsPrint_PrintDocument       
        /// <summary>
        /// Sends an Aspose.Words document to a printer using the XpsPrint API.
        /// </summary>
        /// <param name="document"></param>
        /// <param name="printerName"></param>
        /// <param name="jobName">Job name. Can be null.</param>
        /// <param name="isWait">True to wait for the job to complete. False to return immediately after submitting the job.</param>
        /// <exception cref="Exception">Thrown if any error occurs.</exception>
        public static void Print(Document document, string printerName, string jobName, bool isWait)
        {
            Console.WriteLine("Print");
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            // Use Aspose.Words to convert the document to XPS and store it in a memory stream.
            MemoryStream stream = new MemoryStream();
            document.Save(stream, SaveFormat.Xps);

            stream.Position = 0;
            Console.WriteLine("Saved as Xps");
            Print(stream, printerName, jobName, isWait);
            Console.WriteLine("After Print");
        }
        //ExEnd:XpsPrint_PrintDocument

        //ExStart:XpsPrint_PrintStream        
        /// <summary>
        /// Sends a stream that contains a document in the XPS format to a printer using the XpsPrint API.
        /// Has no dependency on Aspose.Words, can be used in any project.
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="printerName"></param>
        /// <param name="jobName">Job name. Can be null.</param>
        /// <param name="isWait">True to wait for the job to complete. False to return immediately after submitting the job.</param>
        /// <exception cref="Exception">Thrown if any error occurs.</exception>
        public static void Print(Stream stream, string printerName, string jobName, bool isWait)
        {
            if (stream == null)
                throw new ArgumentNullException(nameof(stream));
            if (printerName == null)
                throw new ArgumentNullException(nameof(printerName));

            // Create an event that we will wait on until the job is complete.
            IntPtr completionEvent = CreateEvent(IntPtr.Zero, true, false, null);
            if (completionEvent == IntPtr.Zero)
                throw new Win32Exception();

            Console.WriteLine("StartJob");
            StartJob(printerName, jobName, completionEvent, out IXpsPrintJob job, out IXpsPrintJobStream jobStream);
            Console.WriteLine("Done StartJob");

            Console.WriteLine("Start CopyJob");
            CopyJob(stream, jobStream);
            Console.WriteLine("End CopyJob");

            Console.WriteLine("Start Wait");
            if (isWait)
            {
                WaitForJob(completionEvent);
                CheckJobStatus(job);
            }
            Console.WriteLine("End Wait");

            if (completionEvent != IntPtr.Zero)
                CloseHandle(completionEvent);
            Console.WriteLine("Close Handle");
        }
        //ExEnd:XpsPrint_PrintStream

        private static void StartJob(string printerName, string jobName, IntPtr completionEvent, out IXpsPrintJob job,
            out IXpsPrintJobStream jobStream)
        {
            int result = StartXpsPrintJob(printerName, jobName, null, IntPtr.Zero, completionEvent,
                null, 0, out job, out jobStream, IntPtr.Zero);
            if (result != 0)
                throw new Win32Exception(result);
        }

        private static void CopyJob(Stream stream, IXpsPrintJobStream jobStream)
        {
            byte[] buff = new byte[4096];
            while (true)
            {
                uint read = (uint)stream.Read(buff, 0, buff.Length);
                if (read == 0)
                    break;

                jobStream.Write(buff, read, out uint written);

                if (read != written)
                    throw new Exception("Failed to copy data to the print job stream.");
            }

            // Indicate that the entire document has been copied.
            jobStream.Close();
        }

        private static void WaitForJob(IntPtr completionEvent)
        {
            const int infinite = -1;
            switch (WaitForSingleObject(completionEvent, infinite))
            {
                case WAIT_RESULT.WAIT_OBJECT_0:
                    // Expected result, do nothing.
                    break;
                case WAIT_RESULT.WAIT_FAILED:
                    throw new Win32Exception();
                default:
                    throw new Exception("Unexpected result when waiting for the print job.");
            }
        }

        private static void CheckJobStatus(IXpsPrintJob job)
        {
            job.GetJobStatus(out XPS_JOB_STATUS jobStatus);
            switch (jobStatus.completion)
            {
                case XPS_JOB_COMPLETION.XPS_JOB_COMPLETED:
                    // Expected result, do nothing.
                    break;
                case XPS_JOB_COMPLETION.XPS_JOB_FAILED:
                    throw new Win32Exception(jobStatus.jobStatus);
                default:
                    throw new Exception("Unexpected print job status.");
            }
        }

        [DllImport("XpsPrint.dll", EntryPoint = "StartXpsPrintJob")]
        private static extern int StartXpsPrintJob(
            [MarshalAs(UnmanagedType.LPWStr)] string printerName,
            [MarshalAs(UnmanagedType.LPWStr)] string jobName,
            [MarshalAs(UnmanagedType.LPWStr)] string outputFileName,
            IntPtr progressEvent,
            IntPtr completionEvent,
            [MarshalAs(UnmanagedType.LPArray)] byte[] printablePagesOn,
            uint printablePagesOnCount,
            out IXpsPrintJob xpsPrintJob,
            out IXpsPrintJobStream documentStream,
            IntPtr printTicketStream); // "out IXpsPrintJobStream", we don't use it and just want to pass null, hence IntPtr.

        [DllImport("Kernel32.dll", SetLastError = true)]
        private static extern IntPtr CreateEvent(IntPtr lpEventAttributes, bool bManualReset, bool bInitialState,
            string lpName);

        [DllImport("Kernel32.dll", SetLastError = true, ExactSpelling = true)]
        private static extern WAIT_RESULT WaitForSingleObject(IntPtr handle, int milliseconds);

        [DllImport("Kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool CloseHandle(IntPtr hObject);
    }

    /// <summary>
    /// This interface definition is HACKED.
    /// 
    /// It appears that the IID for IXpsPrintJobStream specified in XpsPrint.h as 
    /// MIDL_INTERFACE("7a77dc5f-45d6-4dff-9307-d8cb846347ca") is not correct and the RCW cannot return it.
    /// But the returned object returns the parent ISequentialStream inteface successfully.
    /// 
    /// So the hack is that we obtain the ISequentialStream interface but work with it as 
    /// with the IXpsPrintJobStream interface. 
    /// </summary>
    [Guid("0C733A30-2A1C-11CE-ADE5-00AA0044773D")] // This is IID of ISequenatialSteam.
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IXpsPrintJobStream
    {
        // ISequentualStream methods.
        void Read([MarshalAs(UnmanagedType.LPArray)] byte[] pv, uint cb, out uint pcbRead);

        void Write([MarshalAs(UnmanagedType.LPArray)] byte[] pv, uint cb, out uint pcbWritten);

        // IXpsPrintJobStream methods.
        void Close();
    }

    [Guid("5ab89b06-8194-425f-ab3b-d7a96e350161")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IXpsPrintJob
    {
        void Cancel();
        void GetJobStatus(out XPS_JOB_STATUS jobStatus);
    }

    [StructLayout(LayoutKind.Sequential)]
    struct XPS_JOB_STATUS
    {
        public uint jobId;
        public int currentDocument;
        public int currentPage;
        public int currentPageTotal;
        public XPS_JOB_COMPLETION completion;
        public int jobStatus;
    };

    enum XPS_JOB_COMPLETION
    {
        XPS_JOB_IN_PROGRESS = 0,
        XPS_JOB_COMPLETED = 1,
        XPS_JOB_CANCELLED = 2,
        XPS_JOB_FAILED = 3
    }

    enum WAIT_RESULT
    {
        WAIT_OBJECT_0 = 0,
        WAIT_ABANDONED = 0x80,
        WAIT_TIMEOUT = 0x102,
        WAIT_FAILED = -1
    }

    //ExStart:ActivePrintPreviewDialogClass 
    internal class ActivePrintPreviewDialog : PrintPreviewDialog
    {
        /// <summary>
        /// Brings the Print Preview dialog on top when it is initially displayed.
        /// </summary>
        protected override void OnShown(EventArgs e)
        {
            Activate();
            base.OnShown(e);
        }
    }
    //ExEnd:ActivePrintPreviewDialogClass
}
#endif