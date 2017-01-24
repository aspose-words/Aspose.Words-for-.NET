using Aspose.Words;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace _07._02_DeleteRange
{
    class Program
    {
        static void Main(string[] args)
        {
            // Check for license and apply if exists
            string licenseFile = AppDomain.CurrentDomain.BaseDirectory + "Aspose.Words.lic";
            if (File.Exists(licenseFile))
            {
                // Apply Aspose.Words API License
                Aspose.Words.License license = new Aspose.Words.License();
                // Place license file in Bin/Debug/ Folder
                license.SetLicense("Aspose.Words.lic");
            }

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Start building a the table.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Row 1, Cell 1 Content");

            // Build the second cell
            builder.InsertCell();
            builder.Write("Row 1, Cell 2 Content");

            // End previous row and start new
            builder.EndRow();

            // Build the first cell of 2nd row
            builder.InsertCell();
            builder.Write("Row 2, Cell 1 Content");

            builder.InsertCell();
            builder.Write("Row 2, Cell 2 Content");

            builder.EndRow();

            // End the table
            builder.EndTable();

            Range range = doc.Sections[0].Range;
            range.Delete();

            String text = doc.Range.Text;

            System.Console.WriteLine(text);
            System.Console.ReadKey();
        }
    }
}
