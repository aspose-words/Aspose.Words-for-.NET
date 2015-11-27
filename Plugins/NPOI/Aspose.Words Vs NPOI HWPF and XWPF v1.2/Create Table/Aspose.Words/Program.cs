using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Aspose.Words
{
    class Program
    {
        static void Main(string[] args)
        {
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

            doc.Save("Aspose_CreateTable.doc");
        }
    }
}
