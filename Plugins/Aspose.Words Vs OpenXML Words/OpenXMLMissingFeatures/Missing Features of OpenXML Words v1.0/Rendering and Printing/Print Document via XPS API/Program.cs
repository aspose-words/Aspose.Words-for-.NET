// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using System;
using System.IO;
using System.Reflection;

namespace XpsPrint
{
    /// <summary>
    /// This sample shows how to convert a document to XPS by means of Aspose.Words and then print with the XpsPrint API.
    /// This sample supports both x86 and x64 platforms.
    /// 
    /// The way to print documents suggested by Microsoft is to use the XpsPrint API 
    /// http://msdn.microsoft.com/en-us/library/dd374565(VS.85).aspx. This API is available on Windows 7, 
    /// Windows Server 2008 R2 and also Windows Vista, provided the Platform Update for Windows Vista is installed.
    /// Since Aspose.Words can easily convert any document into XPS, you can use the following code to print
    /// that document via the XpsPrint API.
    /// </summary>
    class Program
    {
        /// <summary>
        /// The main entry point of the application.
        /// </summary>
        [STAThread]
        public static void Main(string[] args)
        {
            try
            {
                // Sample infrastructure.
                string exeDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + Path.DirectorySeparatorChar;
                string dataDir = new Uri(new Uri(exeDir), @"../Data/").LocalPath;
                //ExStart
                //ExId:XpsPrint_Main
                //ExSummary:Invoke the utility class to print via XPS.
                // Open a sample document in Aspose.Words.
                Aspose.Words.Document document = new Aspose.Words.Document(dataDir + "Print via XPS API.doc");

                // Specify the name of the printer you want to print to.
                const string printerName = @"\\COMPANY\Zeeshan MFC-885CW Printer";

                // Print the document.
                XpsPrintHelper.Print(document, printerName, "test", true);
                //ExEnd
                Console.WriteLine("Printed successfully.");
            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }

            Console.WriteLine("Press Enter.");
            Console.ReadLine();
        }
    }
}
