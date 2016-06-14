
using System;
using System.Collections;
using System.IO;
using System.Drawing;
using Aspose.Words;
using Aspose.Words.Tables;
using System.Diagnostics;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Tables
{
    class InsertTableFromHtml
    {
        public static void Run()
        {
            //ExStart:InsertTableFromHtml
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithTables();
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert the table from HTML. Note that AutoFitSettings does not apply to tables
            // inserted from HTML.
            builder.InsertHtml("<table>" +
                               "<tr>" +
                               "<td>Row 1, Cell 1</td>" +
                               "<td>Row 1, Cell 2</td>" +
                               "</tr>" +
                               "<tr>" +
                               "<td>Row 2, Cell 2</td>" +
                               "<td>Row 2, Cell 2</td>" +
                               "</tr>" +
                               "</table>");

            dataDir = dataDir + "DocumentBuilder.InsertTableFromHtml_out_.doc";
            // Save the document to disk.
            doc.Save(dataDir);
            //ExEnd:InsertTableFromHtml

            Console.WriteLine("\nTable inserted successfully from html.\nFile saved at " + dataDir);
        }
        
    }
}
