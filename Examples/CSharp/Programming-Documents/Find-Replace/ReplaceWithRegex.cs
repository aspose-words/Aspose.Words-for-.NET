using System.Collections;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Fields;
using System;
using Aspose.Words.Replacing;
using NUnit.Framework;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Find_and_Replace
{
    class ReplaceWithRegex
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_FindAndReplace();

            RecognizeAndSubstitutionsWithinReplacementPatterns(dataDir);
            FindAndReplaceWithRegex(dataDir);
        }

        public static void FindAndReplaceWithRegex(string dataDir)
        {
            // ExStart:ReplaceWithRegex
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("sad mad bad");

            Assert.AreEqual("sad mad bad", doc.GetText().Trim());

            // Replaces all occurrences of the words "sad" or "mad" to "bad".
            doc.Range.Replace(new Regex("[s|m]ad"), "bad");

            // Save the Word document.
            doc.Save("Range.ReplaceWithRegex.docx");
            // ExEnd:ReplaceWithRegex
        }
        public static void RecognizeAndSubstitutionsWithinReplacementPatterns(string dataDir)
        {
            // ExStart:RecognizeAndSubstitutionsWithinReplacementPatterns
            // Create new document.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Write some text.
            builder.Write("Jason give money to Paul.");

            Regex regex = new Regex(@"([A-z]+) give money to ([A-z]+)");

            // Replace text using substitutions.
            FindReplaceOptions options = new FindReplaceOptions();
            options.UseSubstitutions = true;
            doc.Range.Replace(regex, @"$2 take money from $1", options);
            // ExEnd:RecognizeAndSubstitutionsWithinReplacementPatterns
            Console.WriteLine(doc.GetText()); // The output is: Paul take money from Jason.\f
        }
    }    
}
