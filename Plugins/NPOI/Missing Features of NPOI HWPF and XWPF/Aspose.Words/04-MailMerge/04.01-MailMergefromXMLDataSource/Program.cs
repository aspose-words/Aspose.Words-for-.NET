using System;
using Aspose.Words;
using System.IO;
using System.Data;

namespace _04._01_MailMergefromXMLDataSource
{
    class Program
    {
        static void Main(string[] args)
        {
            // Check for an Aspose.Words license file in the local file system and apply it, if it exists.
            string licenseFile = AppDomain.CurrentDomain.BaseDirectory + "Aspose.Words.lic";
            if (File.Exists(licenseFile))
            {
                Aspose.Words.License license = new Aspose.Words.License();

                // Use the license from the bin/debug/ Folder.
                license.SetLicense("Aspose.Words.lic");
            }

            // Create the Dataset and read the XML.
            DataSet customersDs = new DataSet();
            customersDs.ReadXml("../../data/Customers.xml");

            // Open a template document.
            Document doc = new Document("../../data/TestFile XML.doc");

            // Execute mail merge to fill the template with data from XML using DataTable.
            doc.MailMerge.Execute(customersDs.Tables["Customer"]);

            // Save the output document.
            doc.Save("MailMergefromXMLDataSource.docx");
        }
    }
}
