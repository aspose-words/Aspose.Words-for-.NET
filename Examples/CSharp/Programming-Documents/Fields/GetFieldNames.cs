using System;
using System.Collections.Generic;
using Aspose.Words;
using Aspose.Words.MailMerging;

namespace CSharp.Programming_Documents.Working_with_Fields
{
    class GetFieldNames
    {
        public static void Run()
        {
            //ExStart:GetFieldNames
            Document doc = new Document();            
            // Shows how to get names of all merge fields in a document.
            string[] fieldNames = doc.MailMerge.GetFieldNames();
            //ExEnd:GetFieldNames
            Console.WriteLine("\nDocument have " + fieldNames.Length + " fields.");
        }
    }
}
