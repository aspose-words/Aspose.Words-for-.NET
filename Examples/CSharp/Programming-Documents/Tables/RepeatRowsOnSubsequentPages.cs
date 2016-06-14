
using System;
using System.Collections;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Diagnostics;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Tables
{
    class RepeatRowsOnSubsequentPages
    {
        public static void Run()
        {
            //ExStart:RepeatRowsOnSubsequentPages
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithTables();

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Table table = builder.StartTable();
            builder.RowFormat.HeadingFormat = true;
            builder.ParagraphFormat.Alignment = ParagraphAlignment.Center;
            builder.CellFormat.Width = 100;
            builder.InsertCell();
            builder.Writeln("Heading row 1");
            builder.EndRow();
            builder.InsertCell();
            builder.Writeln("Heading row 2");
            builder.EndRow();

            builder.CellFormat.Width = 50;
            builder.ParagraphFormat.ClearFormatting();

            // Insert some content so the table is long enough to continue onto the next page.
            for (int i = 0; i < 50; i++)
            {
                builder.InsertCell();
                builder.RowFormat.HeadingFormat = false;
                builder.Write("Column 1 Text");
                builder.InsertCell();
                builder.Write("Column 2 Text");
                builder.EndRow();
            }

            dataDir = dataDir + "Table.HeadingRow_out_.doc";
            // Save the document to disk.
            doc.Save(dataDir);
            //ExEnd:RepeatRowsOnSubsequentPages
            Console.WriteLine("\nTable build successfully which include heading rows that repeat on subsequent pages..\nFile saved at " + dataDir);
        }
        
    }
}
