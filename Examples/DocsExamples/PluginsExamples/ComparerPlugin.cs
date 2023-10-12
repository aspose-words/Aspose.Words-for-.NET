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
            var docA = new Document(MyDir + "Document.docx");
            var docB = new Document(MyDir + "Blank.docx");

            // docA now contains changes as revisions.
            docA.Compare(docB, "User", DateTime.Now, new CompareOptions { IgnoreFormatting = true });

            Console.WriteLine(docA.Revisions.Count == 0 ? "Documents are equal" : "Documents are not equal");
            //ExEnd:CompareDocuments
        }
    }
}
