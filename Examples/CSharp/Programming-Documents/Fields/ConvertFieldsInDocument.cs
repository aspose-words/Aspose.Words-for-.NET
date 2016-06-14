using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class ConvertFieldsInDocument
    {
        public static void Run()
        {
            //ExStart:ConvertFieldsInDocument
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();
            string fileName = "TestFile.doc";
            Document doc = new Document(dataDir + fileName);

            // Pass the appropriate parameters to convert all IF fields encountered in the document (including headers and footers) to static text.
            FieldsHelper.ConvertFieldsToStaticText(doc, FieldType.FieldIf);

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            // Save the document with fields transformed to disk.
            doc.Save(dataDir);
            //ExEnd:ConvertFieldsInDocument
            Console.WriteLine("\nConverted fields to static text in the document successfully.\nFile saved at " + dataDir);
        }
    }
}
