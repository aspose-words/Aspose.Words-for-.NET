using Aspose.Words;
using NUnit.Framework;

namespace DocsExamples.Getting_started
{
    public class Hello_world : DocsExamplesBase
    {
        [Test]
        public void HelloWorld()
        {
            //ExStart:HelloWorld
            //GistDesc:Getting started example
            Document docA = new Document();            
            DocumentBuilder builder = new DocumentBuilder(docA);

            // Insert text to the document start.
            builder.MoveToDocumentStart();
            builder.Write("First Hello World paragraph");

            Document docB = new Document(MyDir + "Document.docx");
            // Add document B to the and of document A, preserving document B formatting.
            docA.AppendDocument(docB, ImportFormatMode.KeepSourceFormatting);
            
            docA.Save(ArtifactsDir + "C:\\Temp\\output_AB.pdf");
            //ExEnd:HelloWorld
        }
    }
}
