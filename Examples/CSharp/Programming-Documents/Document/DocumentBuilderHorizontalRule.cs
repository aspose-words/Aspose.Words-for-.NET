using Aspose.Words.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class DocumentBuilderHorizontalRule
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();

            DocumentBuilderInsertHorizontalRule(dataDir);
            DocumentBuilderHorizontalRuleFormat(dataDir);
        }
        public static void DocumentBuilderInsertHorizontalRule(string dataDir)
        {
            // ExStart:DocumentBuilderInsertHorizontalRule
            // Initialize document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Insert a horizontal rule shape into the document.");
            builder.InsertHorizontalRule();

            dataDir = dataDir + "DocumentBuilder.InsertHorizontalRule_out.doc";
            doc.Save(dataDir);
            // ExEnd:DocumentBuilderInsertHorizontalRule
            Console.WriteLine("\nHorizontal rule is inserted into document successfully.\nFile saved at " + dataDir);
        }

        public static void DocumentBuilderHorizontalRuleFormat(string dataDir)
        {
            // ExStart:DocumentBuilderHorizontalRuleFormat
            DocumentBuilder builder = new DocumentBuilder();

            Shape shape = builder.InsertHorizontalRule();
            HorizontalRuleFormat horizontalRuleFormat = shape.HorizontalRuleFormat;

            horizontalRuleFormat.Alignment = HorizontalRuleAlignment.Center;
            horizontalRuleFormat.WidthPercent = 70;
            horizontalRuleFormat.Height = 3;
            horizontalRuleFormat.Color = Color.Blue;
            horizontalRuleFormat.NoShade = true;

            builder.Document.Save("HorizontalRuleFormat.docx");
            // ExEnd:DocumentBuilderHorizontalRuleFormat
            Console.WriteLine("\nHorizontal rule format inserted into document successfully.\nFile saved at " + dataDir);
        }
    }
}
