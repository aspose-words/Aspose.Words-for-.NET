using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class ConvertFieldsInBody
    {
        public static void Run()
        {
            //ExStart:ConvertFieldsInBody
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();

            string fileName = "TestFile.doc";
            Document doc = new Document(dataDir + fileName);

            // Pass the appropriate parameters to convert PAGE fields encountered to static text only in the body of the first section.
            FieldsHelper.ConvertFieldsToStaticText(doc.FirstSection.Body, FieldType.FieldPage);

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            // Save the document with fields transformed to disk.
            doc.Save(dataDir);
            //ExEnd:ConvertFieldsInBody
            Console.WriteLine("\nConverted fields to static text in the document body successfully.\nFile saved at " + dataDir);
        }
    }
}
