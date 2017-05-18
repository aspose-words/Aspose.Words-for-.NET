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
            Paragraph para = (Paragraph)doc.GetChildNodes(NodeType.Paragraph, true)[1];

            // We want to insert an INCLUDETEXT field like this:
            // { INCLUDETEXT  "c:\\temp\\input.txt" }

            // Create instance of FieldAsk class and lets build the above field code
            FieldInclude fieldinclude = (FieldInclude)para.AppendField(FieldType.FieldInclude, false);
            fieldinclude.BookmarkName = "bookmark";
            fieldinclude.SourceFullName = @"c:\temp\input.docx";

            doc.FirstSection.Body.AppendChild(para);

            // Finally update this TOA field
            fieldinclude.Update();

            dataDir = dataDir + "InsertIncludeFieldWithoutDocumentBuilder_out.doc";
            doc.Save(dataDir);

            // ExEnd:InsertIncludeFieldWithoutDocumentBuilder
        }
    }
}
