using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class InsertMailMergeAddressBlockFieldUsingDOM
    {
        public static void Run()
        {
            //ExStart:InsertMailMergeAddressBlockFieldUsingDOM
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();
            Document doc = new Document(dataDir + "in.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Get paragraph you want to append this merge field to
            Paragraph para = (Paragraph)doc.GetChildNodes(NodeType.Paragraph, true)[1];

            // Move cursor to this paragraph
            builder.MoveTo(para);

            // We want to insert a mail merge address block like this:
            // { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\" }

            // Create instance of FieldAddressBlock class and lets build the above field code
            FieldAddressBlock field = (FieldAddressBlock)builder.InsertField(FieldType.FieldAddressBlock, false);

            // { ADDRESSBLOCK \\c 1" }
            field.IncludeCountryOrRegionName = "1";

            // { ADDRESSBLOCK \\c 1 \\d" }
            field.FormatAddressOnCountryOrRegion = true;

            // { ADDRESSBLOCK \\c 1 \\d \\e Test2 }
            field.ExcludedCountryOrRegionName = "Test2";

            // { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 }
            field.NameAndAddressFormat = "Test3";

            // { ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\" }
            field.LanguageId = "Test 4";

            // Finally update this merge field
            field.Update();

            dataDir = dataDir + "InsertMailMergeAddressBlockFieldUsingDOM_out_.doc";
            doc.Save(dataDir);

            //ExEnd:InsertMailMergeAddressBlockFieldUsingDOM
            Console.WriteLine("\nMail Merge address block field using DOM inserted successfully.\nFile saved at " + dataDir);
        }
    }
}
