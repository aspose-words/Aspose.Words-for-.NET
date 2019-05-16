using Aspose.Words.Fields;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class FieldDisplayResults
    {
        public static void Run()
        {
            // ExStart:FieldDisplayResults
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();

            Document document = new Document(dataDir + "Document.docx");
            document.UpdateFields();

            foreach (Field field in document.Range.Fields)
            {
                Console.WriteLine(field.DisplayResult);
            }
            // ExEnd:FieldDisplayResults
        }
    }
}
