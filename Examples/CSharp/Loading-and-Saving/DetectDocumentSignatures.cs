
using System.IO;
using Aspose.Words;
using System;

namespace Aspose.Words.Examples.CSharp.Loading_Saving
{
    class DetectDocumentSignatures
    {
        public static void Run()
        {
            //ExStart:DetectDocumentSignatures
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_LoadingAndSaving();

            // The path to the document which is to be processed.
            string filePath = dataDir + "Document.Signed.docx";

            FileFormatInfo info = FileFormatUtil.DetectFileFormat(filePath);
            if (info.HasDigitalSignature)
            {
                Console.WriteLine(string.Format("Document {0} has digital signatures, they will be lost if you open/save this document with Aspose.Words.", Path.GetFileName(filePath)));
            }
            //ExEnd:DetectDocumentSignatures            
        }
    }
}
