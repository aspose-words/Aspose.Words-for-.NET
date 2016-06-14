using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class InsertNestedFields
    {
        public static void Run()
        {
            //ExStart:InsertNestedFields
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a few page breaks (just for testing)
            for (int i = 0; i < 5; i++)
                builder.InsertBreak(BreakType.PageBreak);

            // Move the DocumentBuilder cursor into the primary footer.
            builder.MoveToHeaderFooter(HeaderFooterType.FooterPrimary);

            // We want to insert a field like this:
            // { IF {PAGE} <> {NUMPAGES} "See Next Page" "Last Page" }
            Field field = builder.InsertField(@"IF ");
            builder.MoveTo(field.Separator);
            builder.InsertField("PAGE");
            builder.Write(" <> ");
            builder.InsertField("NUMPAGES");
            builder.Write(" \"See Next Page\" \"Last Page\" ");

            // Finally update the outer field to recalcaluate the final value. Doing this will automatically update
            // the inner fields at the same time.
            field.Update();
            dataDir = dataDir + "InsertNestedFields_out_.docx";
            doc.Save(dataDir);
            //ExEnd:InsertNestedFields
            Console.WriteLine("\nInserted nested fields in the document successfully.\nFile saved at " + dataDir);
        }
    }
}
