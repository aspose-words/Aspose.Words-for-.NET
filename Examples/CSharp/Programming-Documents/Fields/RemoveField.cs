using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;

namespace CSharp.Programming_Documents.Working_with_Fields
{
    class RemoveField
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();

            Document doc = new Document(dataDir + "Field.RemoveField.doc");

            Field field = doc.Range.Fields[0];
            // Calling this method completely removes the field from the document.
            field.Remove();

            Console.WriteLine("\nRemoved field from the document successfully.");
        }
    }
}
