
using System.IO;

using Aspose.Words;
using System;

namespace CSharp.Quick_Start
{
    class SimpleMailMerge
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_QuickStart();
            string fileName = "MailMerge Template.doc";

            Document doc = new Document(dataDir + fileName);

            // Fill the fields in the document with user data.
            doc.MailMerge.Execute(
                new string[] { "FullName", "Company", "Address", "Address2", "City" },
                new object[] { "James Bond", "MI5 Headquarters", "Milbank", "", "London" });

            // Saves the document to disk.
            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            doc.Save(dataDir);

            Console.WriteLine("\nMail merge performed successfully.\nFile saved at " + dataDir);
        }
    }
}
