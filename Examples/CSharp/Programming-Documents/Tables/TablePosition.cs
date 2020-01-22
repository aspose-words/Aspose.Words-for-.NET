using Aspose.Words.Drawing;
using Aspose.Words.Tables;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Tables
{
    class TablePosition
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithTables();
            GetTablePosition(dataDir);
            GetFloatingTablePosition(dataDir);
            SetFloatingTablePosition(dataDir);
        }

        private static void GetTablePosition(string dataDir)
        {
            // ExStart:GetTablePosition
            Document doc = new Document(dataDir + "Table.Document.doc");

            // Retrieve the first table in the document.
            Table table = (Table)doc.GetChild(NodeType.Table, 0, true);

            if (table.TextWrapping == TextWrapping.Around)
            {
                Console.WriteLine(table.RelativeHorizontalAlignment);
                Console.WriteLine(table.RelativeVerticalAlignment);
            }
            else
            {
                Console.WriteLine(table.Alignment);
            }

            // ExEnd:GetTablePosition
            Console.WriteLine("\nGet the Table position successfully.");
        }

        private static void GetFloatingTablePosition(string dataDir)
        {
            // ExStart:GetFloatingTablePosition
            Document doc = new Document(dataDir + "FloatingTablePosition.docx");
            foreach (Table table in doc.FirstSection.Body.Tables)
            {
                // If table is floating type then print its positioning properties.
                if (table.TextWrapping == TextWrapping.Around)
                {
                    Console.WriteLine(table.HorizontalAnchor);
                    Console.WriteLine(table.VerticalAnchor);
                    Console.WriteLine(table.AbsoluteHorizontalDistance);
                    Console.WriteLine(table.AbsoluteVerticalDistance);
                    Console.WriteLine(table.AllowOverlap);
                    Console.WriteLine(table.AbsoluteHorizontalDistance);
                    Console.WriteLine(table.RelativeVerticalAlignment);
                    Console.WriteLine("..............................");
                }
            }

            // ExEnd:GetFloatingTablePosition
            Console.WriteLine("\nGet the Table position successfully.");
        }

        private static void SetFloatingTablePosition(string dataDir)
        {
            // ExStart:SetFloatingTablePosition
            Document doc = new Document(dataDir + "FloatingTablePosition.docx");

            Table table = doc.FirstSection.Body.Tables[0];

            // sets absolute table horizontal position at 10pt.
            table.AbsoluteHorizontalDistance = 10;

            // sets vertical table position to center of entity specified by Table.VerticalAnchor.
            table.RelativeVerticalAlignment = VerticalAlignment.Center;

            // Save the document to disk.
            doc.Save(dataDir + "Table.SetFloatingTablePosition.docx");
            // ExEnd:SetFloatingTablePosition
            Console.WriteLine("\nSet the Table position successfully.");
        }
    }
}
