using Aspose.Words.Replacing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Find_and_Replace
{
    class ReplaceHtmlTextWithMeta_Characters
    {
        public static void Run()
        {
            // ExStart:ReplaceHtmlTextWithMetaCharacters
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_FindAndReplace();

            string html = @"<p>&ldquo;Some Text&rdquo;</p>";

            // Initialize a Document.
            Document doc = new Document();

            // Use a document builder to add content to the document.
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Write("{PLACEHOLDER}");

            var findReplaceOptions = new FindReplaceOptions
            {
                ReplacingCallback = new FindAndInsertHtml(),
                PreserveMetaCharacters = true
            };

            doc.Range.Replace("{PLACEHOLDER}", html, findReplaceOptions);


            dataDir = dataDir + "ReplaceHtmlTextWithMetaCharacters_out.doc";
            doc.Save(dataDir);
            // ExEnd:ReplaceHtmlTextWithMetaCharacters
            Console.WriteLine("\nText replaced with meta characters successfully.\nFile saved at " + dataDir);
        }
    }

    // ExStart:ReplaceHtmlFindAndInsertHtml
    public class FindAndInsertHtml : IReplacingCallback
    {
        ReplaceAction IReplacingCallback.Replacing(ReplacingArgs e)
        {
            // This is a Run node that contains either the beginning or the complete match.
            Node currentNode = e.MatchNode;

            // create Document Buidler and insert MergeField
            DocumentBuilder builder = new DocumentBuilder(e.MatchNode.Document as Document);

            builder.MoveTo(currentNode);

            builder.InsertHtml(e.Replacement);

            currentNode.Remove();

            //Signal to the replace engine to do nothing because we have already done all what we wanted.
            return ReplaceAction.Skip;
        }
    }
    // ExEnd:ReplaceHtmlFindAndInsertHtml
}
