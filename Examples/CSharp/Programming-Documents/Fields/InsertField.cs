using System;
using System.Collections;
using System.IO;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Fields;
using Aspose.Words.Layout;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_with_Fields
{
    class InsertField
    {
        public static void Run()
        {
            //ExStart:InsertField
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithFields();

            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.InsertField(@"MERGEFIELD MyFieldName \* MERGEFORMAT");
            dataDir = dataDir + "InsertField_out_.docx";
            doc.Save(dataDir);
            //ExEnd:InsertField
            Console.WriteLine("\nInserted field in the document successfully.\nFile saved at " + dataDir);
        }
    }
}
