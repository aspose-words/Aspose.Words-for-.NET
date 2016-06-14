using System;
using System.Collections;
using System.IO;

using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class ConvertFieldsInParagraph
    {
        public static void Run()
        {
            //ExStart:ConvertFieldsInParagraph
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();
            string fileName = "TestFile.doc";
            Document doc = new Document(dataDir + fileName);

            // Pass the appropriate parameters to convert all IF fields to static text that are encountered only in the last 
            // paragraph of the document.
            FieldsHelper.ConvertFieldsToStaticText(doc.FirstSection.Body.LastParagraph, FieldType.FieldIf);

            dataDir = dataDir + RunExamples.GetOutputFilePath(fileName);
            // Save the document with fields transformed to disk.
            doc.Save(dataDir);
            //ExEnd:ConvertFieldsInParagraph
            Console.WriteLine("\nConverted fields to static text in the paragraph successfully.\nFile saved at " + dataDir);
        }
    }
}
