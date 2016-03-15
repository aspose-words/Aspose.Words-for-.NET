
using System.IO;
using Aspose.Words;
using System;
using System.Web;

namespace CSharp.Loading_Saving
{
    class SendToClientBrowser
    {
        public static void Run()
        {
            //ExStart:SendToClientBrowser
            HttpResponse Respose = null;
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            Document doc = new Document(dataDir + "Document.doc");

            dataDir = dataDir + "Report_out.doc";
            // If this method overload is causing a compiler error then you are using the Client Profile DLL whereas 
            // the Aspose.Words .NET 2.0 DLL must be used instead.
            doc.Save(Respose, dataDir, ContentDisposition.Inline, null);
            //ExEnd:SendToClientBrowser
            Console.WriteLine("\nDocument send to client browser successfully.\nFile saved at " + dataDir);
        }
    }
}
