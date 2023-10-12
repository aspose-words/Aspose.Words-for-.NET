using Aspose.Words;
using Aspose.Words.Comparing;
using NUnit.Framework;

namespace PluginsExamples
{
    public class ComparerPlugin : PluginsExamplesBase
    {
        [Test]
        public void CompareDocuments()
        {
            //ExStart:CompareDocuments
            //GistId:fe5b5017bb9ffbf8a2619a4d90baff33
            var docA = new Document(MyDir + "Blank.docx");
            var docB = new Document(MyDir + "Document.docx");

            // docA now contains changes as revisions.
            docA.Compare(docB, "User", DateTime.Now, new CompareOptions { IgnoreFormatting = true });

            foreach (Revision revision in docA.Revisions)
            {
                Console.WriteLine("Type: " + revision.RevisionType);
                Console.WriteLine("Author: " + revision.Author);
                Console.WriteLine("Date: " + revision.DateTime);
                Console.WriteLine("Revision text: " + revision.ParentNode.ToString(SaveFormat.Text));                
            }
            //ExEnd:CompareDocuments
        }
    }
}
