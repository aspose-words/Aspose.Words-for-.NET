using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Words.Examples.CSharp.Programming_Documents.Working_With_Document
{
    class OpenTypeFeatures
    {
        public static void Run()
        {
            // ExStart:OpenTypeFeatures
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_WorkingWithDocument();

            // Open a document
            Document doc = new Document(dataDir + "OpenType.Document.docx");

            // When text shaper factory is set, layout starts to use OpenType features.
            // An Instance property returns static BasicTextShaperCache object wrapping HarfBuzzTextShaperFactory
            doc.LayoutOptions.TextShaperFactory = Shaping.HarfBuzz.HarfBuzzTextShaperFactory.Instance;

            // Render the document to PDF format
            doc.Save(dataDir + "OpenType.Document.pdf");
            // ExEnd:OpenTypeFeatures
            Console.WriteLine("\nRendered the document with OpenType Features using HarfBuzz shaping.");
        }
    }
}
