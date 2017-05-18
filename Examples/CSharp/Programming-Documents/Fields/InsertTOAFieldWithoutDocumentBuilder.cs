using Aspose.Words.Fields;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class InsertTOAFieldWithoutDocumentBuilder
    {
        public static void Run()
        {
            // ExStart:InsertTOAFieldWithoutDocumentBuilder
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();
            Document doc = new Document(dataDir + "in.doc");
            // Get paragraph you want to append this TOA field to
            Paragraph para = (Paragraph)doc.GetChildNodes(NodeType.Paragraph, true)[1];

            // We want to insert TA and TOA fields like this:
            // { TA  \c 1 \l "Value 0" }
            // { TOA  \c 1 }

            // Create instance of FieldAsk class and lets build the above field code
            FieldTA fieldTA = (FieldTA)para.AppendField(FieldType.FieldTOAEntry, false);
            fieldTA.EntryCategory = "1";
            fieldTA.LongCitation = "Value 0";

            doc.FirstSection.Body.AppendChild(para);

            para = new Paragraph(doc);

            // Create instance of FieldToa class
            FieldToa fieldToa = (FieldToa)para.AppendField(FieldType.FieldTOA, false);
            fieldToa.EntryCategory = "1";
            doc.FirstSection.Body.AppendChild(para);

            // Finally update this TOA field
            fieldToa.Update();

            dataDir = dataDir + "InsertTOAFieldWithoutDocumentBuilder_out.doc";
            doc.Save(dataDir);

            // ExEnd:InsertTOAFieldWithoutDocumentBuilder
            Console.WriteLine("\nTOA field without using document builder inserted successfully.\nFile saved at " + dataDir);
        }
    }
}
