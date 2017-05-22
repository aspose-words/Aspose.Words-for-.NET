using Aspose.Words.Fields;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class InsertFieldIncludeTextWithoutDocumentBuilder
    {
        public static void Run()
        {
            // ExStart:InsertFieldIncludeTextWithoutDocumentBuilder
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();
            Document doc = new Document(dataDir + "in.doc");
            // Get paragraph you want to append this INCLUDETEXT field to
            Paragraph para = (Paragraph)doc.GetChildNodes(NodeType.Paragraph, true)[1];

            // We want to insert an INCLUDETEXT field like this:
            // { INCLUDETEXT  "file path" }

            // Create instance of FieldAsk class and lets build the above field code
            FieldIncludeText fieldIncludeText = (FieldIncludeText)para.AppendField(FieldType.FieldIncludeText, false);
            fieldIncludeText.BookmarkName = "bookmark";
            fieldIncludeText.SourceFullName = dataDir + @"IncludeText.docx";

            doc.FirstSection.Body.AppendChild(para);

            // Finally update this IncludeText field
            fieldIncludeText.Update();

            dataDir = dataDir + "InsertIncludeFieldWithoutDocumentBuilder_out.doc";
            doc.Save(dataDir);

            // ExEnd:InsertFieldIncludeTextWithoutDocumentBuilder
            Console.WriteLine("\nIncludeText field without using document builder inserted successfully.\nFile saved at " + dataDir);
        }
    }
}
