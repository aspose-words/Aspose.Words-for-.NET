
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
            CompareDocumentWithComparisonTarget(dataDir);
            SpecifyComparisonGranularity(dataDir);
        }             

        private static void NormalComparison(string dataDir)
        {
            // ExStart:NormalComparison
            Document docA = new Document(dataDir + "TestFile.doc");
            Document docB = new Document(dataDir + "TestFile - Copy.doc");
            // DocA now contains changes as revisions. 
            docA.Compare(docB, "user", DateTime.Now); 
            // ExEnd:NormalComparison                     
        }
        private static void CompareForEqual(string dataDir)
        {
            // ExStart:CompareForEqual
            Document docA = new Document(dataDir + "TestFile.doc");
            Document docB = new Document(dataDir + "TestFile - Copy.doc");
            // DocA now contains changes as revisions. 
            docA.Compare(docB, "user", DateTime.Now);
            if (docA.Revisions.Count == 0)
                Console.WriteLine("Documents are equal");
            else
                Console.WriteLine("Documents are not equal");
            // ExEnd:CompareForEqual                     
        }

        private static void CompareDocumentWithCompareOptions(string dataDir)
        {
            // ExStart:CompareDocumentWithCompareOptions
            Document docA = new Document(dataDir + "TestFile.doc");
            Document docB = new Document(dataDir + "TestFile - Copy.doc");

            CompareOptions options = new CompareOptions();
            options.IgnoreFormatting = true;
            options.IgnoreHeadersAndFooters = true;
            options.IgnoreCaseChanges = true;
            options.IgnoreTables = true;
            options.IgnoreFields = true;
            options.IgnoreComments = true;
            options.IgnoreTextboxes = true;
            options.IgnoreFootnotes = true;

            // DocA now contains changes as revisions. 
            docA.Compare(docB, "user", DateTime.Now, options);
            if (docA.Revisions.Count == 0)
                Console.WriteLine("Documents are equal");
            else
                Console.WriteLine("Documents are not equal");
            // ExEnd:CompareDocumentWithCompareOptions                     
        }

        private static void CompareDocumentWithComparisonTarget(string dataDir)
        {
            // ExStart:CompareDocumentWithComparisonTarget
            Document docA = new Document(dataDir + "TestFile.doc");
            Document docB = new Document(dataDir + "TestFile - Copy.doc");

            CompareOptions options = new CompareOptions();
            options.IgnoreFormatting = true;
            // Relates to Microsoft Word "Show changes in" option in "Compare Documents" dialog box. 
            options.Target = ComparisonTargetType.New;

            docA.Compare(docB, "user", DateTime.Now, options);
            // ExEnd:CompareDocumentWithComparisonTarget      

            dataDir = dataDir + "TestFile_Out.doc";

            Console.WriteLine("\nDocuments have compared successfully.\nFile saved at " + dataDir);
        }

        public static void SpecifyComparisonGranularity(string dataDir)
        {
            // ExStart:SpecifyComparisonGranularity
            DocumentBuilder builderA = new DocumentBuilder(new Document());
            DocumentBuilder builderB = new DocumentBuilder(new Document());

            builderA.Writeln("This is A simple word");
            builderB.Writeln("This is B simple words");

            CompareOptions co = new CompareOptions();
            co.Granularity = Granularity.CharLevel;

            builderA.Document.Compare(builderB.Document, "author", DateTime.Now, co);
            // ExEnd:SpecifyComparisonGranularity      
        }
    }
}
