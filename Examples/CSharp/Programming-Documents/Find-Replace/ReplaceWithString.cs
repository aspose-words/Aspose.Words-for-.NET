using System.Collections;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Words;
using Aspose.Words.Fields;
using System;
using System.Drawing;
using Aspose.Words.Replacing;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Find_and_Replace
{
    class ReplaceWithString
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_FindAndReplace();

            Replace(dataDir);
            HighlightColor(dataDir);
        }

        public static void Replace(string dataDir)
        {
            // ExStart:ReplaceWithString
            // Load a Word Docx document by creating an instance of the Document class.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello _CustomerName_,");

            // Specify the search string and replace string using the Replace method.
            doc.Range.Replace("_CustomerName_", "James Bond", new FindReplaceOptions());

            // Save the result.
            doc.Save(dataDir + "Range.ReplaceSimple.docx");
            // ExEnd:ReplaceWithString
        }

        public static void HighlightColor(string dataDir)
        {
            // Load a Word Docx document by creating an instance of the Document class.
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            builder.Writeln("Hello _CustomerName_,");

            // ExStart:HighlightColor
            // Highlight word "the" with yellow color.
            FindReplaceOptions options = new FindReplaceOptions();
            options.ApplyFont.HighlightColor = Color.Yellow;

            // Replace highlighted text.
            doc.Range.Replace("Hello", "Hello", options);
            // ExEnd:HighlightColor
        }
    }

    
}
