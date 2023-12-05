using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Getting_started
{
    public class HelloWorld : DocsExamplesBase
    {
        [Test]
        public void SimpleHelloWorld()
        {
            //ExStart:HelloWorld
            //GistId:542a463e1857480986d18ec296ed43d5
            Document docA = new Document();            
            DocumentBuilder builder = new DocumentBuilder(docA);

            // Insert text to the document start.
            builder.MoveToDocumentStart();
            builder.Write("First Hello World paragraph");

            Document docB = new Document(MyDir + "Document.docx");
            // Add document B to the and of document A, preserving document B formatting.
            docA.AppendDocument(docB, ImportFormatMode.KeepSourceFormatting);
            
            docA.Save(ArtifactsDir + "HelloWorld.SimpleHelloWorld.pdf");
            //ExEnd:HelloWorld
        }
    }
}
