using Aspose.Words.Fields;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class InsertFieldNone
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();
            // ExStart:InsertFieldNone
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            FieldUnknown field = (FieldUnknown)builder.InsertField(FieldType.FieldNone, false);

            dataDir = dataDir + "InsertFieldNone_out.docx";
            doc.Save(dataDir);
            // ExEnd:InsertFieldNone
            Console.WriteLine("\nInserted field in the document successfully.\nFile saved at " + dataDir);
        }
    }
}
