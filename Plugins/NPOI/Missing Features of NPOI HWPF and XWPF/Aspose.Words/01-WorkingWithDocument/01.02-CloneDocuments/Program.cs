using System;
using System.Collections.Generic;
using System.Text; using Aspose.Words;

namespace _01._02_CloneDocuments
{
    class Program
    {
        static void Main(string[] args)
        {
            Document doc = new Document("../../data/document.doc");
            Document clone = doc.Clone();

            clone.Save("AsposeClone.doc", SaveFormat.Doc);
        }
    }
}
