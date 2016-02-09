
using System.IO;
using Aspose.Words;
using System;

namespace CSharp.Programming_Documents.Working_With_Document
{
    class ProtectDocument
    {
        public static void Run()
        {            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            dataDir = dataDir + "ProtectDocument.doc";

            Protect(dataDir);
            UnProtect(dataDir);
            GetProtectionType(dataDir);
        }
        /// <summary>
        /// Shows how to protect document
        /// </summary>
        /// <param name="inputFileName">input file name with complete path.</param>        
        public static void Protect(string inputFileName)
        {
            //ExStart:ProtectDocument
            Document doc = new Document(inputFileName);
            doc.Protect(ProtectionType.AllowOnlyFormFields, "password");
            Console.WriteLine("\nDocument protected successfully.");
            //ExEnd:ProtectDocument
        }
        /// <summary>
        /// Shows how to unprotect document
        /// </summary>
        /// <param name="inputFileName">input file name with complete path.</param>        
        public static void UnProtect(string inputFileName)
        {
            //ExStart:UnProtectDocument
            Document doc = new Document(inputFileName);
            doc.Unprotect();
            //ExEnd:UnProtectDocument
            Console.WriteLine("\nDocument unprotected successfully.");
        }
        /// <summary>
        /// Shows how to get protection type
        /// </summary>
        /// <param name="inputFileName">input file name with complete path.</param>        
        public static void GetProtectionType(string inputFileName)
        {
            //ExStart:GetProtectionType
            Document doc = new Document(inputFileName);
            ProtectionType protectionType = doc.ProtectionType;
            //ExEnd:GetProtectionType
            Console.WriteLine("\nDocument protection type is " + protectionType.ToString());
        }
    }
}
