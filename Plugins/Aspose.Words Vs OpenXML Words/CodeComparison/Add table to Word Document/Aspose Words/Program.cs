// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Words;

namespace Aspose_Words
{
    class Program
    {
        static void Main(string[] args)
        {
            
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // We call this method to start building the table.
            builder.StartTable();
            builder.InsertCell();
            builder.Write("Row 1, Cell 1 Content.");

            // Build the second cell
            builder.InsertCell();
            builder.Write("Row 1, Cell 2 Content.");
            // Call the following method to end the row and start a new row.
            builder.EndRow();

            // Build the first cell of the second row.
            builder.InsertCell();
            builder.Write("Row 2, Cell 1 Content");

            // Build the second cell.
            builder.InsertCell();
            builder.Write("Row 2, Cell 2 Content.");
            builder.EndRow();

            // Signal that we have finished building the table.
            builder.EndTable();

            // Save the document to disk.
            doc.Save("DocumentBuilder.CreateSimpleTable Out.doc");
        }
    }
}
