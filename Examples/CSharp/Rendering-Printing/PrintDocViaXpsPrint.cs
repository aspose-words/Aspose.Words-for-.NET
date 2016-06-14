using System;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Layout;
using Aspose.Words.Rendering;
using Aspose.Words.Saving;
using Aspose.Words.Drawing;
using System.Drawing.Imaging;
using Aspose.Words.Tables;
using System.ComponentModel;
using System.Runtime.InteropServices;
namespace Aspose.Words.Examples.CSharp.Rendering_and_Printing
{
    class PrintDocViaXpsPrint
    {
        public static void Run()
        {
            //ExStart:PrintDocViaXpsPrint
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_RenderingAndPrinting();
            // Open a sample document in Aspose.Words.
            Document document = new Document(dataDir + "TestFile.doc");

            // Specify the name of the printer you want to print to.
            const string printerName = @"\\COMPANY\Brother MFC-885CW Printer";

            // Print the document.
            XpsPrintHelper.Print(document, printerName, "My Test Job", true);
            //ExEnd:PrintDocViaXpsPrint
        }
        
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
        //ExSummary:Convert an Aspose.Words document into an XPS stream and print.
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
                throw new ArgumentNullException("document");

            // Use Aspose.Words to convert the document to XPS and store in a memory stream.
            MemoryStream stream = new MemoryStream();
            document.Save(stream, SaveFormat.Xps);

            stream.Position = 0;
            Console.WriteLine("Saved as Xps");
            Print(stream, printerName, jobName, isWait);
            Console.WriteLine("After Print");
        }
        //ExEnd:XpsPrint_PrintDocument
        //ExStart:XpsPrint_PrintStream        
        //ExSummary:Prints an XPS document using the XpsPrint API.
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
                throw new ArgumentNullException("stream");
            if (printerName == null)
                throw new ArgumentNullException("printerName");

            // Create an event that we will wait on until the job is complete.
            IntPtr completionEvent = CreateEvent(IntPtr.Zero, true, false, null);
            if (completionEvent == IntPtr.Zero)
                throw new Win32Exception();

            //            try
            //            {
            IXpsPrintJob job;
            IXpsPrintJobStream jobStream;
            Console.WriteLine("StartJob");
            StartJob(printerName, jobName, completionEvent, out job, out jobStream);
            Console.WriteLine("Done StartJob");
            Console.WriteLine("Start CopyJob");
            CopyJob(stream, job, jobStream);
            Console.WriteLine("End CopyJob");

            Console.WriteLine("Start Wait");
            if (isWait)
            {
                WaitForJob(completionEvent);
                CheckJobStatus(job);
            }
            Console.WriteLine("End Wait");
            /*            }
                        finally
                        {
                            if (completionEvent != IntPtr.Zero)
                                CloseHandle(completionEvent);
                        }
            */
            if (completionEvent != IntPtr.Zero)
                CloseHandle(completionEvent);
            Console.WriteLine("Close Handle");
        }
        //ExEnd:XpsPrint_PrintStream

        private static void StartJob(string printerName, string jobName, IntPtr completionEvent, out IXpsPrintJob job, out IXpsPrintJobStream jobStream)
        {
            int result = StartXpsPrintJob(printerName, jobName, null, IntPtr.Zero, completionEvent,
                null, 0, out job, out jobStream, IntPtr.Zero);
            if (result != 0)
                throw new Win32Exception(result);
        }

        private static void CopyJob(Stream stream, IXpsPrintJob job, IXpsPrintJobStream jobStream)
        {

            //            try
            //            {
            byte[] buff = new byte[4096];
            while (true)
            {
                uint read = (uint)stream.Read(buff, 0, buff.Length);
                if (read == 0)
                    break;

                uint written;
                jobStream.Write(buff, read, out written);

                if (read != written)
                    throw new Exception("Failed to copy data to the print job stream.");
            }

            // Indicate that the entire document has been copied.
            jobStream.Close();
            //            }
            //            catch (Exception)
            //            {
            //                // Cancel the job if we had any trouble submitting it.
            //                job.Cancel();
            //                throw;
            //            }
        }

        private static void WaitForJob(IntPtr completionEvent)
        {
            const int INFINITE = -1;
            switch (WaitForSingleObject(completionEvent, INFINITE))
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
            XPS_JOB_STATUS jobStatus;
            job.GetJobStatus(out jobStatus);
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
            [MarshalAs(UnmanagedType.LPWStr)] String printerName,
            [MarshalAs(UnmanagedType.LPWStr)] String jobName,
            [MarshalAs(UnmanagedType.LPWStr)] String outputFileName,
            IntPtr progressEvent,   // HANDLE
            IntPtr completionEvent, // HANDLE
            [MarshalAs(UnmanagedType.LPArray)] byte[] printablePagesOn,
            UInt32 printablePagesOnCount,
            out IXpsPrintJob xpsPrintJob,
            out IXpsPrintJobStream documentStream,
            IntPtr printTicketStream);  // This is actually "out IXpsPrintJobStream", but we don't use it and just want to pass null, hence IntPtr.

        [DllImport("Kernel32.dll", SetLastError = true)]
        private static extern IntPtr CreateEvent(IntPtr lpEventAttributes, bool bManualReset, bool bInitialState, string lpName);

        [DllImport("Kernel32.dll", SetLastError = true, ExactSpelling = true)]
        private static extern WAIT_RESULT WaitForSingleObject(IntPtr handle, Int32 milliseconds);

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
    [Guid("0C733A30-2A1C-11CE-ADE5-00AA0044773D")]  // This is IID of ISequenatialSteam.
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IXpsPrintJobStream
    {
        // ISequentualStream methods.
        void Read([MarshalAs(UnmanagedType.LPArray)] byte[] pv, uint cb, out uint pcbRead);
        void Write([MarshalAs(UnmanagedType.LPArray)] byte[] pv, uint cb, out uint pcbWritten);
        // IXpsPrintJobStream methods.
        void Close();
    }

    [Guid("5ab89b06-8194-425f-ab3b-d7a96e350161")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IXpsPrintJob
    {
        void Cancel();
        void GetJobStatus(out XPS_JOB_STATUS jobStatus);
    }

    [StructLayout(LayoutKind.Sequential)]
    struct XPS_JOB_STATUS
    {
        public UInt32 jobId;
        public Int32 currentDocument;
        public Int32 currentPage;
        public Int32 currentPageTotal;
        public XPS_JOB_COMPLETION completion;
        public Int32 jobStatus; // UInt32
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
        WAIT_FAILED = -1 // 0xFFFFFFFF
    }
}
