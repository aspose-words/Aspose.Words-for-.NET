
using System.IO;
using Aspose.Words;
using System;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class CompareDocument
    {
        public static void Run()
        {            
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();
            NormalComparison(dataDir);
            CompareForEqual(dataDir);
        }             
        private static void NormalComparison(string dataDir)
        {
            //ExStart:NormalComparison
            Document docA = new Document(dataDir + "TestFile.doc");
            Document docB = new Document(dataDir + "TestFile - Copy.doc");
            // docA now contains changes as revisions. 
            docA.Compare(docB, "user", DateTime.Now); 
            //ExEnd:NormalComparison                     
        }
        private static void CompareForEqual(string dataDir)
        {
            //ExStart:CompareForEqual
            Document docA = new Document(dataDir + "TestFile.doc");
            Document docB = new Document(dataDir + "TestFile - Copy.doc");
            // docA now contains changes as revisions. 
            docA.Compare(docB, "user", DateTime.Now);
            if (docA.Revisions.Count == 0)
                Console.WriteLine("Documents are equal");
            else
                Console.WriteLine("Documents are not equal");
            //ExEnd:CompareForEqual                     
        }  
    }
}
