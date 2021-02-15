using Aspose.Words;
using System;
using System.IO;

namespace _07._01_AccessRanges
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

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building the table.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Row 1, Cell 1 Content");

            // Build the second cell.
            builder.InsertCell();
            builder.Write("Row 1, Cell 2 Content");

            // End previous row and start new.
            builder.EndRow();

            // Build the first cell of 2nd row.
            builder.InsertCell();
            builder.Write("Row 2, Cell 1 Content");

            builder.InsertCell();
            builder.Write("Row 2, Cell 2 Content");

            builder.EndRow();

            // End the table.
            builder.EndTable();

            Range range = doc.Range;
            string text = range.Text;

            Console.WriteLine(text);
            Console.ReadKey();
        }
    }
}
