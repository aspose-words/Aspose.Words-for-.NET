using Aspose.Words.Fields;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class InsertIncludeFieldWithoutDocumentBuilder
    {
        public static void Run()
        {
            // ExStart:InsertIncludeFieldWithoutDocumentBuilder
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();
            Document doc = new Document(dataDir + "in.doc");
            // Get paragraph you want to append this INCLUDETEXT field to
            Paragraph para = (Paragraph)doc.GetChildNodes(NodeType.Paragraph, true)[1];

            // We want to insert an INCLUDETEXT field like this:
            // { INCLUDETEXT  "file path" }

            // Create instance of FieldAsk class and lets build the above field code
            FieldInclude fieldinclude = (FieldInclude)para.AppendField(FieldType.FieldInclude, false);
            fieldinclude.BookmarkName = "bookmark";
            fieldinclude.SourceFullName = dataDir + @"IncludeText.docx";

            doc.FirstSection.Body.AppendChild(para);

            // Finally update this Include field
            fieldinclude.Update();

            dataDir = dataDir + "InsertIncludeFieldWithoutDocumentBuilder_out.doc";
            doc.Save(dataDir);

            // ExEnd:InsertIncludeFieldWithoutDocumentBuilder
            Console.WriteLine("\nInclude field without using document builder inserted successfully.\nFile saved at " + dataDir);
        }
    }
}
