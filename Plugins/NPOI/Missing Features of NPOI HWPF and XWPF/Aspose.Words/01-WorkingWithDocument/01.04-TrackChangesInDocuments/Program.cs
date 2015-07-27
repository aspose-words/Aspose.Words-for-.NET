using System;
using System.Collections.Generic;
using System.Text; using Aspose.Words;

namespace _01._04_TrackChangesInDocuments
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document("../../data/document.doc");
            doc.AcceptAllRevisions();

            doc.Save("AsposeAcceptChanges.doc", SaveFormat.Doc);
        }
    }
}
