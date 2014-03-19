//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Words;
using Aspose.Words.Fields;

namespace InsertNestedFieldsExample
{
    public class Program
    {
        public static void Main()
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //ExStart
            //ExFor:DocumentBuilder.InsertField(string)
            //ExId:DocumentBuilderInsertNestedFields
            //ExSummary:Demonstrates how to insert fields nested within another field using DocumentBuilder.
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

            doc.Save(dataDir + "InsertNestedFields Out.docx");
            //ExEnd		
        }
    }
}