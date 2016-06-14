using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class InsertMergeFieldUsingDOM
    {
        public static void Run()
        {
            //ExStart:InsertMergeFieldUsingDOM
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();
            Document doc = new Document(dataDir + "in.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Get paragraph you want to append this merge field to
            Paragraph para = (Paragraph)doc.GetChildNodes(NodeType.Paragraph, true)[1];

            // Move cursor to this paragraph
            builder.MoveTo(para);

            // We want to insert a merge field like this:
            // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v" }

            // Create instance of FieldMergeField class and lets build the above field code
            FieldMergeField field = (FieldMergeField)builder.InsertField(FieldType.FieldMergeField, false);

            // { " MERGEFIELD Test1" }
            field.FieldName = "Test1";

            // { " MERGEFIELD Test1 \\b Test2" }
            field.TextBefore = "Test2";

            // { " MERGEFIELD Test1 \\b Test2 \\f Test3 }
            field.TextAfter = "Test3";

            // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m" }
            field.IsMapped = true;

            // { " MERGEFIELD Test1 \\b Test2 \\f Test3 \\m \\v" }
            field.IsVerticalFormatting = true;

            // Finally update this merge field
            field.Update();

            dataDir = dataDir + "InsertMergeFieldUsingDOM_out_.doc";
            doc.Save(dataDir);
            
            //ExEnd:InsertMergeFieldUsingDOM
            Console.WriteLine("\nMerge field using DOM inserted successfully.\nFile saved at " + dataDir);
        }
    }
}
