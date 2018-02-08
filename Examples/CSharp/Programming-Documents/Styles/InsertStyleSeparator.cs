using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Styles
{
    class InsertStyleSeparator
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithStyles();

            // ExStart:ParagraphInsertStyleSeparator
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Style paraStyle = builder.Document.Styles.Add(StyleType.Paragraph, "MyParaStyle");
            paraStyle.Font.Bold = false;
            paraStyle.Font.Size = 8;
            paraStyle.Font.Name = "Arial";

            // Append text with "Heading 1" style.
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;
            builder.Write("Heading 1");
            builder.InsertStyleSeparator();

            // Append text with another style.
            builder.ParagraphFormat.StyleName = paraStyle.Name;
            builder.Write("This is text with some other formatting ");
             
            dataDir = dataDir + "BindSDTtoCustomXmlPart_out.doc";

            // Save the document to disk.
            doc.Save(dataDir);
            // ExEnd:ParagraphInsertStyleSeparator 
            
            Console.WriteLine("\nApplied different paragraph styles to two different parts of a text line successfully.");
        }
    }
}
