// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

//ExStart
//ExId:MossDoc2Pdf
//ExSummary:The following is the complete code of the document converter.

using System;
using System.IO;

using Aspose.Words;
using Aspose.Words.Saving;

namespace ApiExamples
{
    /// <summary>
    /// DOC2PDF document converter for SharePoint.
    /// Uses Aspose.Words to perform the conversion.
    /// </summary>
    public class ExMossDoc2Pdf
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main(string[] args)
        {
            // Although SharePoint passes "-log <filename>" to us and we are
            // supposed to log there, for the sake of simplicity, we will use 
            // our own hard coded path to the log file.
            // 
            // Make sure there are permissions to write into this folder.
            // The document converter will be called under the document 
            // conversion account (not sure what name), so for testing purposes 
            // I would give the Users group write permissions into this folder.
            gLog = new StreamWriter(@"C:\Aspose2Pdf\log.txt", true);

            try
            {
                gLog.WriteLine(DateTime.Now.ToString() + " Started");
                gLog.WriteLine(Environment.CommandLine);

                ParseCommandLine(args);

                // Uncomment the code below when you have purchased a licenses for Aspose.Words.
                //
                // You need to deploy the license in the same folder as your 
                // executable, alternatively you can add the license file as an 
                // embedded resource to your project.
                //
                // // Set license for Aspose.Words.
                // Aspose.Words.License wordsLicense = new Aspose.Words.License();
                // wordsLicense.SetLicense("Aspose.Total.lic");

                ConvertDoc2Pdf(gInFileName, gOutFileName);
            }
            catch (Exception e)
            {
                gLog.WriteLine(e.Message);
                Environment.ExitCode = 100;
            }
            finally
            {
                gLog.Close();
            }
        }

        private static void ParseCommandLine(string[] args)
        {
            int i = 0;
            while (i < args.Length)
            {
                string s = args[i];
                switch (s.ToLower())
                {
                    case "-in":
                        i++;
                        gInFileName = args[i];
                        break;
                    case "-out":
                        i++;
                        gOutFileName = args[i];
                        break;
                    case "-config":
                        // Skip the name of the config file and do nothing.
                        i++;
                        break;
                    case "-log":
                        // Skip the name of the log file and do nothing.
                        i++;
                        break;
                    default:
                        throw new Exception("Unknown command line argument: " + s);
                }
                i++;
            }
        }

        private static void ConvertDoc2Pdf(string inFileName, string outFileName)
        {
            // You can load not only DOC here, but any format supported by
            // Aspose.Words: DOC, DOCX, RTF, WordML, HTML, MHTML, ODT etc.
            Document doc = new Document(inFileName);

            doc.Save(outFileName, new PdfSaveOptions());
        }

        private static string gInFileName;
        private static string gOutFileName;
        private static StreamWriter gLog;
    }
}
//ExEnd


